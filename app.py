import logging
import os
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urlparse

import requests
from dotenv import load_dotenv
from flask import Flask, flash, redirect, render_template, request, send_from_directory, url_for
from openpyxl import Workbook
from pymongo import MongoClient
from pymongo.errors import PyMongoError


BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "local-development-key")
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

PICTURES_DIR = BASE_DIR / "Online_Pictures"
DOWNLOADS_DIR = BASE_DIR / "downloads"
PICTURES_DIR.mkdir(exist_ok=True)
DOWNLOADS_DIR.mkdir(exist_ok=True)

CACHE_TTL_MINUTES = max(1, int(os.getenv("CACHE_TTL_MINUTES", "60")))
SERPAPI_URL = "https://serpapi.com/search.json"
PRODUCT_FIELDS = [
    "Product_Name",
    "Product_Price",
    "Commission_fee",
    "Total_price",
    "global_rating",
    "retailer_name",
    "Retailer_email",
    "Product_picture_name",
]
EXCEL_HEADERS = [
    "Product Name",
    "Product Price (USD)",
    "Commission Fee (USD)",
    "Total Price (USD)",
    "Global Rating",
    "Retailer Name",
    "Retailer Email",
    "Product Picture Name",
]


def get_collections():
    uri = os.getenv("MONGODB_URI", "").strip()
    if not uri:
        raise RuntimeError("MongoDB is not configured. Set MONGODB_URI in your .env file.")
    client = MongoClient(uri, serverSelectionTimeoutMS=3000)
    client.admin.command("ping")
    database = client[os.getenv("MONGODB_DATABASE", "smartcart")]
    return client, database


def safe_filename(value):
    cleaned = re.sub(r"[^A-Za-z0-9_-]+", "_", value.strip()).strip("_")
    return cleaned or "student"


def export_filename(searched_at=None):
    name_parts = [
        os.getenv("STUDENT_FIRST_NAME", "YourFirstName"),
        os.getenv("STUDENT_LAST_NAME", "YourLastName"),
        os.getenv("REGISTRATION_NUMBER", "YourRegistrationNumber"),
    ]
    timestamp = (searched_at or datetime.now().astimezone()).strftime(
        "%Y%m%d_%H%M%S_%f"
    )
    return (
        "_".join(safe_filename(part) for part in name_parts)
        + f"_{timestamp}.xlsx"
    )


def download_picture(url, asin):
    if not url:
        return "Not available"
    parsed = urlparse(url)
    if parsed.scheme != "https" or parsed.hostname not in {
        "m.media-amazon.com",
        "images-na.ssl-images-amazon.com",
    }:
        logger.warning("Skipped product image from an unexpected host")
        return "Not available"

    filename = f"{safe_filename(asin or 'product')}.jpg"
    destination = PICTURES_DIR / filename
    if destination.exists():
        return filename

    try:
        response = requests.get(url, timeout=12, stream=True)
        response.raise_for_status()
        with destination.open("wb") as image_file:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    image_file.write(chunk)
        return filename
    except requests.RequestException:
        logger.warning("Could not download product image for %s", asin)
        return "Not available"


def search_amazon(term):
    api_key = os.getenv("SERPAPI_API_KEY", "").strip()
    if not api_key or api_key == "add_your_private_key_here":
        raise RuntimeError(
            "Add your SerpApi key to the SERPAPI_API_KEY setting in your local .env file."
        )
    try:
        response = requests.get(
            SERPAPI_URL,
            params={
                "engine": "amazon",
                "amazon_domain": "amazon.com",
                "k": term,
                "api_key": api_key,
            },
            timeout=25,
        )
        response.raise_for_status()
        payload = response.json()
    except requests.RequestException as error:
        raise RuntimeError(f"The Amazon product search service could not be reached: {error}") from error
    except ValueError as error:
        raise RuntimeError("The product search service returned an invalid response.") from error

    if payload.get("error"):
        raise RuntimeError(f"Product search failed: {payload['error']}")

    products = []
    for item in payload.get("organic_results", []):
        price = item.get("extracted_price")
        if price is None:
            continue
        try:
            price = float(price)
        except (TypeError, ValueError):
            continue

        commission = round(price * 0.10, 2)
        products.append(
            {
                "Product_Name": item.get("title") or "Not available",
                "Product_Price": round(price, 2),
                "Commission_fee": commission,
                "Total_price": round(price + commission, 2),
                "global_rating": item.get("rating", "Not available"),
                "retailer_name": "Amazon",
                "Retailer_email": "Not available",
                "Product_picture_name": download_picture(
                    item.get("thumbnail"), item.get("asin")
                ),
            }
        )
    return products


def desktop_directory():
    desktop = Path.home() / "Desktop"
    if not desktop.is_dir():
        onedrive_desktop = Path.home() / "OneDrive" / "Desktop"
        if onedrive_desktop.is_dir():
            return onedrive_desktop
        raise RuntimeError(
            f"Desktop folder was not found at {desktop}. Create it or update the "
            "desktop_directory setting in app.py."
        )
    return desktop


def write_excel(products):
    filename = export_filename()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "SmartCart Products"
    sheet.append(EXCEL_HEADERS)
    for product in products:
        sheet.append([product.get(field, "Not available") for field in PRODUCT_FIELDS])
    workbook.save(DOWNLOADS_DIR / filename)
    workbook.save(desktop_directory() / filename)
    return filename


def record_search(database, term, products, source):
    searched_at = datetime.now(timezone.utc)
    database.search_history.insert_one(
        {
            "search_term": term,
            "timestamp": searched_at,
            "number_of_results": len(products),
            "source": source,
        }
    )
    database.commission_log.insert_one(
        {
            "search_term": term,
            "timestamp": searched_at,
            "products": products,
            "total_commission": round(
                sum(product["Commission_fee"] for product in products), 2
            ),
            "result_count": len(products),
            "source": source,
        }
    )


def find_products(term):
    client, database = get_collections()
    try:
        cutoff = datetime.now(timezone.utc) - timedelta(minutes=CACHE_TTL_MINUTES)
        cached = database.product_cache.find_one(
            {"search_term": term.casefold(), "fetched_at": {"$gte": cutoff}},
            sort=[("fetched_at", -1)],
        )
        if cached:
            products = cached["products"]
            source = "cache"
        else:
            products = search_amazon(term)
            source = "live"
            database.product_cache.insert_one(
                {
                    "search_term": term.casefold(),
                    "fetched_at": datetime.now(timezone.utc),
                    "products": products,
                }
            )
        record_search(database, term, products, source)
        return products, source
    finally:
        client.close()


@app.get("/")
def index():
    return render_template("index.html", products=None, query="", excel_filename=None)


@app.post("/search")
def search():
    term = request.form.get("product_name", "").strip()
    if not term:
        flash("Enter a product name to search.", "error")
        return redirect(url_for("index"))
    if len(term) > 150:
        flash("Product names must be 150 characters or fewer.", "error")
        return redirect(url_for("index"))

    try:
        products, source = find_products(term)
        excel_filename = write_excel(products)
    except (RuntimeError, PyMongoError, OSError) as error:
        logger.error("Search could not be completed: %s", error)
        flash(str(error), "error")
        return render_template(
            "index.html", products=None, query=term, excel_filename=None
        ), 503

    if not products:
        flash("No priced products were found for that search.", "notice")
    return render_template(
        "index.html",
        products=products,
        query=term,
        source=source,
        excel_filename=excel_filename,
    )


@app.get("/downloads/<path:filename>")
def download_excel(filename):
    return send_from_directory(DOWNLOADS_DIR, filename, as_attachment=True)


@app.get("/pictures/<path:filename>")
def product_picture(filename):
    return send_from_directory(PICTURES_DIR, filename)


if __name__ == "__main__":
    app.run(debug=True)

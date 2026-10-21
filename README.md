# SmartCart

A Flask school project that searches Amazon.com through SerpApi, adds a 10% commission, keeps a MongoDB cache and history, and exports results to Excel.

## What you need

- Python 3.10 or newer
- MongoDB Community Server running locally, or a MongoDB connection URI
- A SerpApi account. The current Free plan lists 250 searches per month; check their current terms and limits before use.

SerpApi's Amazon engine returns current listing data such as title, price, rating, and image URL. The search response does not provide a retailer email, so SmartCart displays `Not available` rather than inventing one. Search-result pages also return a limited list of results, not every Amazon listing.

## Set up on Windows

Open PowerShell in this folder and run:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .env.example .env
```

Open `.env` in VS Code or Notepad. Replace `add_your_private_key_here` with your key from the SerpApi dashboard. Enter your first name, last name, and registration number in the three student settings. The key is read by the Python app through `python-dotenv`; `.env` is excluded from Git and should never be shared or committed. Keep MongoDB running.

Start the app:

```powershell
python app.py
```

Open `http://127.0.0.1:5000` in your browser.

## Where files and data go

- MongoDB database `smartcart`: `product_cache`, `search_history`, and `commission_log`.
- Downloaded product images: `Online_Pictures`.
- Generated Excel workbook: every search creates a separate file, automatically saved to the Desktop of the computer running this local Flask app and copied into `downloads` for the page's download link. Each filename starts with `YourFirstName_LastName_RegistrationNumber` and adds the search timestamp (including microseconds) to prevent later searches from overwriting earlier files. For this direct Desktop save to work as intended, run Flask on your own computer. If Windows redirects Desktop to OneDrive, the app checks `OneDrive\Desktop` as well.

The cache lasts 60 minutes by default (`CACHE_TTL_MINUTES`). Every successful search, including a cache hit, creates a search history record and a commission log record. A fresh result list is cached before it is shown.

## Notes

Prices are shown in the currency returned by Amazon.com (usually USD). Commission is calculated from the listing price, rounded to cents. Retailer email is not supplied by the search provider. Individual image download failures do not prevent product search from working. Product pictures are stored under `Online_Pictures` and displayed on each result card when available.

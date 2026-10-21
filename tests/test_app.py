import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from datetime import datetime, timezone

import app


SAMPLE_PRODUCTS = [
    {
        "Product_Name": "Sample headphones",
        "Product_Price": 50.0,
        "Commission_fee": 5.0,
        "Total_price": 55.0,
        "global_rating": 4.5,
        "retailer_name": "Amazon",
        "Retailer_email": "Not available",
        "Product_picture_name": "Not available",
    }
]


class FakeCollection:
    def __init__(self):
        self.documents = []

    def find_one(self, query, sort=None):
        matches = [
            document
            for document in self.documents
            if document["search_term"] == query["search_term"]
            and document["fetched_at"] >= query["fetched_at"]["$gte"]
        ]
        return max(matches, key=lambda document: document["fetched_at"]) if matches else None

    def insert_one(self, document):
        self.documents.append(document)


class FakeDatabase:
    def __init__(self):
        self.product_cache = FakeCollection()
        self.search_history = FakeCollection()
        self.commission_log = FakeCollection()


class FakeClient:
    def __init__(self):
        self.closed = False

    def close(self):
        self.closed = True


class SmartCartTests(unittest.TestCase):
    def setUp(self):
        app.app.config["TESTING"] = True
        self.client = app.app.test_client()

    def test_homepage_renders_search_form(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b'name="product_name"', response.data)

    def test_empty_search_is_rejected(self):
        response = self.client.post("/search", data={"product_name": " "})
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.headers["Location"], "/")

    def test_successful_search_renders_results_and_export_link(self):
        with patch.object(app, "find_products", return_value=(SAMPLE_PRODUCTS, "live")), patch.object(
            app, "write_excel", return_value="Student_Name_Reg.xlsx"
        ):
            response = self.client.post(
                "/search", data={"product_name": "headphones"}
            )
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Sample headphones", response.data)
        self.assertIn(b"product-card", response.data)
        self.assertIn(b"$5.00", response.data)
        self.assertIn(b"$55.00", response.data)
        self.assertIn(b"Student_Name_Reg.xlsx", response.data)

    def test_search_uses_recent_cache_and_logs_each_search(self):
        database = FakeDatabase()
        client = FakeClient()
        with patch.object(app, "get_collections", return_value=(client, database)), patch.object(
            app, "search_amazon", return_value=SAMPLE_PRODUCTS
        ) as live_search:
            first = app.find_products("Headphones")
            second = app.find_products("headphones")

        self.assertEqual(first, (SAMPLE_PRODUCTS, "live"))
        self.assertEqual(second, (SAMPLE_PRODUCTS, "cache"))
        live_search.assert_called_once_with("Headphones")
        self.assertEqual(len(database.product_cache.documents), 1)
        self.assertEqual(len(database.search_history.documents), 2)
        self.assertEqual(len(database.commission_log.documents), 2)
        self.assertEqual(database.commission_log.documents[0]["total_commission"], 5.0)
        self.assertTrue(client.closed)

    def test_export_filename_uses_student_details(self):
        with patch.dict(
            os.environ,
            {
                "STUDENT_FIRST_NAME": "Jane",
                "STUDENT_LAST_NAME": "Doe",
                "REGISTRATION_NUMBER": "24-123",
            },
        ):
            searched_at = datetime(2026, 10, 2, 12, 0, tzinfo=timezone.utc)
            self.assertEqual(
                app.export_filename(searched_at),
                "Jane_Doe_24-123_20261002_120000_000000.xlsx",
            )

    def test_each_export_gets_a_distinct_filename(self):
        with patch.dict(
            os.environ,
            {
                "STUDENT_FIRST_NAME": "Jane",
                "STUDENT_LAST_NAME": "Doe",
                "REGISTRATION_NUMBER": "24-123",
            },
        ):
            first = app.export_filename(
                datetime(2026, 10, 2, 12, 0, 0, 1, tzinfo=timezone.utc)
            )
            second = app.export_filename(
                datetime(2026, 10, 2, 12, 0, 0, 2, tzinfo=timezone.utc)
            )
        self.assertNotEqual(first, second)
        self.assertTrue(first.startswith("Jane_Doe_24-123_"))
        self.assertTrue(second.startswith("Jane_Doe_24-123_"))

    def test_excel_is_saved_to_downloads_and_desktop(self):
        with tempfile.TemporaryDirectory() as temporary_directory:
            downloads = Path(temporary_directory) / "downloads"
            desktop = Path(temporary_directory) / "Desktop"
            downloads.mkdir()
            desktop.mkdir()
            with patch.object(app, "DOWNLOADS_DIR", downloads), patch.object(
                app, "desktop_directory", return_value=desktop
            ), patch.object(app, "export_filename", return_value="Jane_Doe_24.xlsx"):
                filename = app.write_excel(SAMPLE_PRODUCTS)

            self.assertTrue((downloads / filename).is_file())
            self.assertTrue((desktop / filename).is_file())


if __name__ == "__main__":
    unittest.main()

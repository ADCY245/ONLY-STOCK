import unittest
from datetime import datetime

from bson import ObjectId

from backend import app as app_module


class WarehouseIntegrationTests(unittest.TestCase):
    def test_location_id_validation_is_stable_and_normalized(self):
        location_id, error = app_module.parse_location_id("wh-a-rb02-l1-b")
        self.assertIsNone(error)
        self.assertEqual(location_id, "WH-A-RB02-L1-B")

        location_id, error = app_module.parse_location_id("Rack B shelf 1")
        self.assertIsNone(location_id)
        self.assertIn("WH-A-RACK-L1-A", error)

    def test_serialized_inventory_exposes_optional_location_id(self):
        now = datetime.now(app_module.IST_TIMEZONE)
        item = {
            "_id": ObjectId(),
            "category": "Cutting Rules",
            "brand": "Test",
            "type": "coil",
            "batch_roll_no": None,
            "width": None,
            "height": None,
            "size": "",
            "thickness": None,
            "quantity": 2,
            "unit": "coil",
            "location_id": "WH-A-ER01-L3-A",
            "created_at": now,
            "updated_at": now,
        }
        self.assertEqual(app_module.serialize_item(item)["location_id"], "WH-A-ER01-L3-A")

    def test_warehouse_assets_are_served_by_flask(self):
        client = app_module.app.test_client()
        for path in ("/warehouse.js", "/warehouse.css"):
            response = client.get(path)
            try:
                self.assertEqual(response.status_code, 200, path)
            finally:
                response.close()

    def test_frontend_has_one_shared_pdf_derived_location_model(self):
        warehouse_source = (app_module.FRONTEND_DIR / "warehouse.js").read_text(encoding="utf-8")
        index_source = (app_module.FRONTEND_DIR / "index.html").read_text(encoding="utf-8")
        self.assertIn('data-page="warehouse"', index_source)
        self.assertIn("const WAREHOUSE_CONFIG", warehouse_source)
        self.assertIn("const WAREHOUSE_LOCATIONS", warehouse_source)
        self.assertIn('sourceWall: "South wall elevation"', warehouse_source)
        self.assertIn('sourceWall: "North wall elevation"', warehouse_source)
        self.assertIn('sourceWall: "East wall elevation"', warehouse_source)


if __name__ == "__main__":
    unittest.main()

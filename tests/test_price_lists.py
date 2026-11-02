import csv
import re
import unittest
from copy import deepcopy
from io import BytesIO, StringIO
from types import SimpleNamespace
from unittest.mock import patch

from bson import ObjectId
from openpyxl import Workbook, load_workbook
from pymongo.errors import PyMongoError
import xlwt
from werkzeug.datastructures import FileStorage
from werkzeug.exceptions import RequestEntityTooLarge

from backend.app import PRICE_LIST_UPLOAD_MAX_BYTES, _price_list_search_tokens, app
from backend.price_lists import build_price_list_workbook, parse_price_list_file, parse_product_description


def upload_bytes(content, filename):
    return FileStorage(stream=BytesIO(content), filename=filename)


def matches(document, query):
    for key, expected in query.items():
        if key == "$and":
            if not all(matches(document, option) for option in expected):
                return False
            continue
        if key == "$or":
            if not any(matches(document, option) for option in expected):
                return False
            continue
        actual = document.get(key)
        if isinstance(expected, dict):
            if "$regex" in expected:
                if not re.search(expected["$regex"], str(actual or ""), re.IGNORECASE):
                    return False
            if "$in" in expected and actual not in expected["$in"]:
                return False
            if "$ne" in expected and actual == expected["$ne"]:
                return False
        elif actual != expected:
            return False
    return True


class FakeCursor(list):
    def sort(self, fields, direction=None):
        if isinstance(fields, str):
            fields = [(fields, direction or 1)]
        for field, order in reversed(fields):
            super().sort(key=lambda item: item.get(field) or "", reverse=order < 0)
        return self

    def skip(self, count):
        return FakeCursor(self[count:])

    def limit(self, count):
        return FakeCursor(self[:count])


class FakeCollection:
    def __init__(self):
        self.items = []
        self.sequence = 0

    def find_one(self, query, sort=None, session=None):
        items = [item for item in self.items if matches(item, query)]
        if sort:
            items = FakeCursor(items).sort(sort)
        return items[0] if items else None

    def insert_one(self, item, session=None):
        self.sequence += 1
        item.setdefault("_id", self.sequence)
        self.items.append(item.copy())
        return SimpleNamespace(inserted_id=item["_id"])

    def insert_many(self, items, ordered=True):
        for item in items:
            self.insert_one(item)

    def update_many(self, query, update):
        for item in self.items:
            if matches(item, query):
                item.update(update.get("$set", {}))

    def update_one(self, query, update, upsert=False, session=None):
        item = self.find_one(query)
        if item:
            item.update(update.get("$set", {}))
            return
        if upsert:
            new_item = {key: value for key, value in query.items() if not key.startswith("$") and not isinstance(value, dict)}
            new_item.update(update.get("$setOnInsert", {}))
            new_item.update(update.get("$set", {}))
            self.insert_one(new_item)

    def delete_one(self, query):
        self.items = [item for item in self.items if not matches(item, query)]

    def delete_many(self, query):
        self.items = [item for item in self.items if not matches(item, query)]

    def find(self, query=None):
        return FakeCursor([item for item in self.items if matches(item, query or {})])

    def count_documents(self, query, session=None):
        return len(self.find(query))

    def distinct(self, field, query):
        return list({item.get(field) for item in self.find(query)})

    def aggregate(self, pipeline):
        match_stage = pipeline[0]["$match"]
        groups = {}
        for item in self.find(match_stage):
            groups[item["customer_name"]] = groups.get(item["customer_name"], 0) + 1
        return [{"_id": name, "count": count} for name, count in sorted(groups.items())]


class BooleanRejectingCollection(FakeCollection):
    """Mirror PyMongo Collection's deliberate rejection of truth testing."""

    def __bool__(self):
        raise TypeError(
            "Collection objects do not implement truth value testing or bool(). "
            "Please compare with None instead: collection is not None"
        )


class FailingRevisionCollection(FakeCollection):
    def insert_one(self, item, session=None):
        raise PyMongoError("history persistence failed")


class RollbackFakeSession:
    def __init__(self, collections):
        self.collections = collections
        self.snapshots = None

    def start_transaction(self):
        self.snapshots = [(collection, deepcopy(collection.items), collection.sequence) for collection in self.collections]
        return self

    def __enter__(self):
        return self

    def __exit__(self, exc_type, *_):
        if exc_type and self.snapshots:
            for collection, items, sequence in self.snapshots:
                collection.items = items
                collection.sequence = sequence
        return False


class RollbackFakeClient:
    def __init__(self, collections):
        self.collections = collections

    def start_session(self):
        return RollbackFakeSession(self.collections)


class FakeTransaction:
    def start_transaction(self):
        return self

    def __enter__(self):
        return self

    def __exit__(self, *_):
        return False


class FakeClient:
    def start_session(self):
        return FakeTransaction()


class FakeDatabase:
    client = FakeClient()


class PriceListTests(unittest.TestCase):
    def test_price_list_upload_limit_returns_json_413(self):
        self.assertEqual(app.config["MAX_CONTENT_LENGTH"], 25 * 1024 * 1024)
        with app.test_request_context("/price-lists/preview", method="POST"):
            response, status = app.handle_http_exception(RequestEntityTooLarge())
        self.assertEqual(status, 413)
        self.assertEqual(response.get_json()["max_upload_bytes"], PRICE_LIST_UPLOAD_MAX_BYTES)

    def test_universal_price_list_search_token_normalization(self):
        self.assertEqual(_price_list_search_tokens("1070 × 500"), ["1070", "500"])
        self.assertEqual(_price_list_search_tokens("B4P 1070"), ["b4p", "1070"])
        self.assertEqual(_price_list_search_tokens("  KBA---500 Mtr  "), ["kba", "500", "mtr"])

    def test_universal_price_list_search_matches_all_tokens_across_fields_and_scopes(self):
        imports = FakeCollection()
        entries = FakeCollection()
        overrides = FakeCollection()
        import_id = "import-universal-search"
        imports.items.append({"import_id": import_id, "active": True, "filename": "prices.csv", "stats": {}})
        entries.items.extend([
            {
                "_id": ObjectId(), "import_id": import_id, "customer_name": "Customer A",
                "brand": "B4P", "product_type": "KBA 105 - Jumbo Auto Wash Cloth",
                "size_code": "1070 mm x 500 Mtr", "original_item_description": "B4P KBA 105 - Jumbo Auto Wash Cloth - 1070 mm x 500 Mtr",
                "price_list_id": "SKU-44717", "price": 18988.0,
            },
            {
                "_id": ObjectId(), "import_id": import_id, "customer_name": "Customer B",
                "brand": "AMC", "product_type": "Blanket Plotter", "size_code": "1070 mm",
                "original_item_description": "AMC Blanket Plotter - 1070 mm", "price_list_id": "SKU-20000", "price": 20000.0,
            },
            {
                "_id": ObjectId(), "import_id": import_id, "customer_name": "Customer A",
                "brand": "B4P", "product_type": "Alu Bar", "size_code": "44843",
                "original_item_description": "B4P Alu Bar EXSQ 44843", "price_list_id": "SKU-44843", "price": 1.0,
            },
        ])
        user = {"_id": ObjectId(), "email": "staff@chemo.in", "role": "workshop"}
        with (
            patch("backend.app.require_auth", return_value=(user, None)),
            patch("backend.app.get_price_list_imports_collection", return_value=imports),
            patch("backend.app.get_price_list_entries_collection", return_value=entries),
            patch("backend.app.get_price_list_overrides_collection", return_value=overrides),
        ):
            client = app.test_client()
            expected_counts = {
                "B4P": 2, "KBA": 1, "Jumbo": 1, "Auto Wash": 1, "1070": 2,
                "500 Mtr": 1, "B4P Jumbo 1070": 1, "b4p jumbo 1070": 1, "1070 B4P KBA": 1,
                "KBA 500": 1, "44717": 1, "1070 x 500": 1,
                "Customer B": 1,
            }
            for search, expected_count in expected_counts.items():
                response = client.get(f"/price-lists/entries?search={search}")
                self.assertEqual(response.status_code, 200, search)
                self.assertEqual(response.get_json()["total"], expected_count, search)
                if expected_count == 1:
                    expected_customer = "Customer B" if search == "Customer B" else "Customer A"
                    self.assertEqual(response.get_json()["entries"][0]["customer_name"], expected_customer, search)

            self.assertEqual(client.get("/price-lists/entries?search=not-found").get_json()["total"], 0)
            self.assertEqual(client.get("/price-lists/entries?search=").get_json()["total"], 3)
            self.assertEqual(client.get("/price-lists/entries?customer=Customer%20B&search=1070%20AMC").get_json()["total"], 1)
            self.assertEqual(client.get("/price-lists/entries?customer=Customer%20B&search=B4P").get_json()["total"], 0)
            paginated = client.get("/price-lists/entries?search=B4P&page=1&page_size=10").get_json()
            self.assertEqual(paginated["total"], 2)
            self.assertEqual(paginated["page_size"], 10)
            self.assertEqual(len(paginated["entries"]), 2)
    def test_csv_mapping_parser_and_transparent_diagnostics(self):
        content = (
            "PriceList ID,PriceList Name,Item Name,PriceList Rate\n"
            "1, M R TRADING ,Image Web Master GR 750 mm x 29.00 mtr x 1.70 mm,3100\n"
            "1,M R TRADING,Image Web Master GR 750 mm x 29.00 mtr x 1.70 mm,3100\n"
            "2,OTHER,MTech Calx De Glazer 1 L,320\n"
            "3,BROKEN,,not-a-price\n"
        ).encode("utf-8")
        parsed = parse_price_list_file(upload_bytes(content, "prices.csv"))
        self.assertEqual(parsed["stats"]["total_records"], 4)
        self.assertEqual(parsed["stats"]["unique_customers"], 3)
        self.assertEqual(parsed["stats"]["invalid_records"], 1)
        self.assertEqual(parsed["stats"]["duplicate_records"], 1)
        self.assertEqual(parsed["stats"]["successful_import_count"], 2)
        self.assertEqual(parsed["records"][0]["customer_name"], "M R TRADING")
        self.assertEqual(parsed["records"][0]["original_item_description"], "Image Web Master GR 750 mm x 29.00 mtr x 1.70 mm")
        self.assertEqual(parsed["records"][0]["size_code"], "750 mm x 29.00 mtr x 1.70 mm")

    def test_three_column_file_uses_optional_id_mapping(self):
        content = b"Customer,Description,Rate\nA,Product 44717,100\n"
        parsed = parse_price_list_file(upload_bytes(content, "prices.csv"))
        self.assertIsNone(parsed["mapping"]["price_list_id"])
        self.assertEqual(parsed["records"][0]["size_code"], "44717")

        content_with_notes = b"Account,Description,Rate,Notes\nA,Product 44717,100,legacy\n"
        parsed_with_notes = parse_price_list_file(upload_bytes(content_with_notes, "prices.csv"))
        self.assertIsNone(parsed_with_notes["mapping"]["price_list_id"])

    def test_xlsx_mapping_supports_different_headers(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Account", "Description", "Customer Price", "Reference"])
        sheet.append(["Customer A", "MTech Active SF 780 mm x 25 mtr", 4305, "PL-1"])
        output = BytesIO()
        workbook.save(output)
        parsed = parse_price_list_file(upload_bytes(output.getvalue(), "prices.xlsx"))
        self.assertEqual(parsed["stats"]["successful_import_count"], 1)
        self.assertEqual(parsed["mapping"]["customer_name"], "Account")
        self.assertEqual(parsed["mapping"]["price"], "Customer Price")
        self.assertEqual(parsed["records"][0]["price_list_id"], "PL-1")

    def test_legacy_xls_is_supported(self):
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Prices")
        for column, value in enumerate(["Customer", "Item", "Rate"]):
            sheet.write(0, column, value)
        for column, value in enumerate(["Customer A", "MTech Active SF 780 mm", 4305]):
            sheet.write(1, column, value)
        output = BytesIO()
        workbook.save(output)
        parsed = parse_price_list_file(upload_bytes(output.getvalue(), "prices.xls"))
        self.assertEqual(parsed["stats"]["successful_import_count"], 1)

    def test_product_parser_keeps_unknown_descriptions_searchable(self):
        parsed = parse_product_description("Uncatalogued product without dimensions")
        self.assertEqual(parsed["brand"], "Uncatalogued")
        self.assertEqual(parsed["product_type"], "product without dimensions")
        self.assertEqual(parsed["size_code"], None)
        self.assertEqual(parsed["original_item_description"], "Uncatalogued product without dimensions")

    def test_workbook_export_has_customer_sheets_and_excel_formatting(self):
        records = [
            {
                "customer_name": "A Customer / North",
                "brand": "Image",
                "product_type": "Sunrise",
                "size_code": "1070 mm",
                "original_item_description": "Image Sunrise 1070 mm",
                "price": 1234.5,
                "price_list_id": "1",
            },
            {
                "customer_name": "B Customer",
                "brand": "MTech",
                "product_type": "Active SF",
                "size_code": "780 mm",
                "original_item_description": "MTech Active SF 780 mm",
                "price": 4305,
                "price_list_id": "2",
            },
        ]
        workbook = build_price_list_workbook(records)
        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        loaded = load_workbook(output)
        self.assertEqual(len(loaded.sheetnames), 2)
        self.assertTrue(all(len(name) <= 31 for name in loaded.sheetnames))
        self.assertEqual(loaded.worksheets[0].freeze_panes, "A2")
        self.assertEqual(loaded.worksheets[0].auto_filter.ref, "A1:G2")
        self.assertEqual(loaded.worksheets[0]["E2"].number_format, "#,##0.00")
        self.assertEqual(loaded.worksheets[0]["F2"].number_format, "#,##0.00")

    def test_preview_route_does_not_touch_inventory(self):
        content = b"PriceList Name,Item Name,PriceList Rate\nA,Image Sunrise 1070 mm,100\n"
        with patch("backend.app.require_role", return_value=({"role": "workshop"}, None)):
            response = app.test_client().post(
                "/price-lists/preview",
                data={"file": (BytesIO(content), "prices.csv")},
                content_type="multipart/form-data",
            )
        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertEqual(payload["stats"]["successful_import_count"], 1)
        self.assertTrue(payload["can_import"])

    def test_import_lookup_replacement_confirmation_and_export(self):
        content = b"PriceList Name,Item Name,PriceList Rate\nA,Image Sunrise 1070 mm,100\nB,Image Sunrise 1070 mm,110\n"
        imports = FakeCollection()
        entries = FakeCollection()
        overrides = FakeCollection()
        with (
            patch("backend.app.require_role", return_value=({"role": "workshop"}, None)),
            patch("backend.app.require_auth", return_value=({"role": "workshop"}, None)),
            patch("backend.app.get_price_list_imports_collection", return_value=imports),
            patch("backend.app.get_price_list_entries_collection", return_value=entries),
            patch("backend.app.get_price_list_overrides_collection", return_value=overrides),
        ):
            client = app.test_client()
            first = client.post("/price-lists/import", data={"file": (BytesIO(content), "prices.csv")}, content_type="multipart/form-data")
            self.assertEqual(first.status_code, 200)
            second = client.post("/price-lists/import", data={"file": (BytesIO(content), "prices.csv")}, content_type="multipart/form-data")
            self.assertEqual(second.status_code, 409)
            customers = client.get("/price-lists/customers")
            self.assertEqual([item["name"] for item in customers.get_json()["customers"]], ["A", "B"])
            lookup = client.get("/price-lists/entries?search=Sunrise")
            self.assertEqual(lookup.status_code, 200)
            self.assertEqual(lookup.get_json()["total"], 2)
            exported = client.get("/price-lists/export?mode=search&search=Sunrise")
            self.assertEqual(exported.status_code, 200)
            workbook = load_workbook(BytesIO(exported.data), read_only=True)
            self.assertEqual(workbook.sheetnames, ["Search Results"])
            replaced = client.post("/price-lists/import", data={"file": (BytesIO(content), "prices.csv"), "replace_existing": "true"}, content_type="multipart/form-data")
            self.assertEqual(replaced.status_code, 200)
            self.assertEqual(len(imports.items), 2)

    def test_price_revisions_are_immutable_and_restore_creates_new_revision(self):
        content = b"PriceList Name,Item Name,PriceList Rate\nA,Image Sunrise 1070 mm,18988\nB,Image Sunrise 1070 mm,19000\n"
        imports = FakeCollection()
        entries = FakeCollection()
        overrides = FakeCollection()
        revisions = FakeCollection()
        user = {"_id": ObjectId(), "email": "staff@chemo.in", "role": "workshop"}
        with (
            patch("backend.app.require_role", return_value=(user, None)),
            patch("backend.app.require_auth", return_value=(user, None)),
            patch("backend.app.get_database", return_value=FakeDatabase()),
            patch("backend.app.get_price_list_imports_collection", return_value=imports),
            patch("backend.app.get_price_list_entries_collection", return_value=entries),
            patch("backend.app.get_price_list_overrides_collection", return_value=overrides),
            patch("backend.app.get_price_list_revisions_collection", return_value=revisions),
        ):
            client = app.test_client()
            imported = client.post("/price-lists/import", data={"file": (BytesIO(content), "prices.csv")}, content_type="multipart/form-data")
            self.assertEqual(imported.status_code, 200)
            entry_id = ObjectId()
            other_entry_id = ObjectId()
            entries.items[0]["_id"] = entry_id
            entries.items[1]["_id"] = other_entry_id
            entry_id_text = str(entry_id)

            initial = client.get("/price-lists/entries")
            self.assertEqual(initial.get_json()["entries"][0]["original_price"], 18988.0)
            self.assertEqual(initial.get_json()["entries"][0]["updated_price"], 18988.0)
            self.assertEqual(initial.get_json()["entries"][1]["updated_price"], 19000.0)

            first_revision = client.patch(
                f"/price-lists/entries/{entry_id_text}/price",
                json={"updated_price": 19500},
            )
            self.assertEqual(first_revision.status_code, 200)
            self.assertEqual(first_revision.get_json()["revision_number"], 1)
            self.assertEqual(len(revisions.items), 1)

            second_revision = client.patch(
                f"/price-lists/entries/{entry_id_text}/price",
                json={"updated_price": 20000},
            )
            self.assertEqual(second_revision.get_json()["revision_number"], 2)
            self.assertEqual(len(revisions.items), 2)
            self.assertEqual(entries.items[0]["price"], 18988.0)

            third_revision = client.patch(
                f"/price-lists/entries/{entry_id_text}/price",
                json={"updated_price": 19750},
            )
            self.assertEqual(third_revision.status_code, 200)
            self.assertEqual(third_revision.get_json()["revision_number"], 3)
            self.assertEqual(third_revision.get_json()["entry"]["original_price"], 18988.0)
            self.assertEqual(third_revision.get_json()["entry"]["updated_price"], 19750.0)
            self.assertEqual(
                [(item["previous_price"], item["new_price"]) for item in revisions.items],
                [(18988.0, 19500.0), (19500.0, 20000.0), (20000.0, 19750.0)],
            )

            history = client.get(f"/price-lists/entries/{entry_id_text}/history")
            self.assertEqual([item["revision_number"] for item in history.get_json()["revisions"]], [1, 2, 3])

            restored = client.post(
                f"/price-lists/entries/{entry_id_text}/restore",
                json={"revision_number": 1},
            )
            self.assertEqual(restored.status_code, 200)
            self.assertEqual(restored.get_json()["revision_number"], 4)
            self.assertEqual(restored.get_json()["entry"]["original_price"], 18988.0)
            self.assertEqual(restored.get_json()["entry"]["updated_price"], 19500.0)
            self.assertEqual(len(revisions.items), 4)
            other_lookup = client.get("/price-lists/entries?customer=B")
            self.assertEqual(other_lookup.get_json()["entries"][0]["updated_price"], 19000.0)

            history_export = client.get("/price-lists/revisions/export")
            self.assertEqual(history_export.status_code, 200)
            exported = load_workbook(BytesIO(history_export.data), read_only=True)
            self.assertEqual(exported.sheetnames, ["Price Revision History"])

        with patch("backend.app.require_auth", return_value=({"_id": ObjectId(), "email": "user@chemo.in", "role": "user"}, None)):
            forbidden = app.test_client().patch(
                f"/price-lists/entries/{entry_id_text}/price",
                json={"updated_price": 175},
            )
        self.assertEqual(forbidden.status_code, 403)

    def test_price_revision_does_not_truth_test_pymongo_collections(self):
        imports = FakeCollection()
        entries = BooleanRejectingCollection()
        overrides = BooleanRejectingCollection()
        revisions = BooleanRejectingCollection()
        import_id = "import-1"
        entry_id = ObjectId()
        imports.items.append({"import_id": import_id, "active": True, "filename": "prices.csv", "stats": {}})
        entries.items.append(
            {
                "_id": entry_id,
                "import_id": import_id,
                "customer_name": "A",
                "brand": "Image",
                "product_type": "Sunrise",
                "size_code": "1070 mm",
                "original_item_description": "Image Sunrise 1070 mm",
                "price": 18988.0,
            }
        )
        user = {"_id": ObjectId(), "email": "staff@chemo.in", "role": "workshop"}
        with (
            patch("backend.app.require_role", return_value=(user, None)),
            patch("backend.app.get_database", return_value=FakeDatabase()),
            patch("backend.app.get_price_list_imports_collection", return_value=imports),
            patch("backend.app.get_price_list_entries_collection", return_value=entries),
            patch("backend.app.get_price_list_overrides_collection", return_value=overrides),
            patch("backend.app.get_price_list_revisions_collection", return_value=revisions),
        ):
            response = app.test_client().patch(
                f"/price-lists/entries/{entry_id}/price",
                json={"updated_price": 19500},
            )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["revision_number"], 1)
        self.assertEqual(revisions.items[0]["new_price"], 19500.0)

    def test_price_revision_rejects_invalid_numeric_input(self):
        imports = FakeCollection()
        imports.items.append({"import_id": "import-1", "active": True, "filename": "prices.csv", "stats": {}})
        user = {"_id": ObjectId(), "email": "staff@chemo.in", "role": "workshop"}
        with (
            patch("backend.app.require_role", return_value=(user, None)),
            patch("backend.app.get_price_list_imports_collection", return_value=imports),
        ):
            response = app.test_client().patch(
                f"/price-lists/entries/{ObjectId()}/price",
                json={"updated_price": "not-a-number"},
            )

        self.assertEqual(response.status_code, 400)
        self.assertIn("numeric", response.get_json()["error"])

    def test_price_revision_returns_503_when_mongodb_is_unavailable(self):
        imports = FakeCollection()
        entries = FakeCollection()
        overrides = FakeCollection()
        revisions = FakeCollection()
        entry_id = ObjectId()
        imports.items.append({"import_id": "import-1", "active": True, "filename": "prices.csv", "stats": {}})
        entries.items.append({"_id": entry_id, "import_id": "import-1", "price": 18988.0})
        user = {"_id": ObjectId(), "email": "staff@chemo.in", "role": "workshop"}
        with (
            patch("backend.app.require_role", return_value=(user, None)),
            patch("backend.app.get_database", return_value=SimpleNamespace(client=None)),
            patch("backend.app.get_price_list_imports_collection", return_value=imports),
            patch("backend.app.get_price_list_entries_collection", return_value=entries),
            patch("backend.app.get_price_list_overrides_collection", return_value=overrides),
            patch("backend.app.get_price_list_revisions_collection", return_value=revisions),
        ):
            response = app.test_client().patch(
                f"/price-lists/entries/{entry_id}/price",
                json={"updated_price": 19500},
            )

        self.assertEqual(response.status_code, 503)
        self.assertIn("transaction support", response.get_json()["error"])

    def test_failed_history_persistence_rolls_back_override(self):
        imports = FakeCollection()
        entries = FakeCollection()
        overrides = FakeCollection()
        revisions = FailingRevisionCollection()
        entry_id = ObjectId()
        imports.items.append({"import_id": "import-1", "active": True, "filename": "prices.csv", "stats": {}})
        entries.items.append({"_id": entry_id, "import_id": "import-1", "price": 18988.0})
        user = {"_id": ObjectId(), "email": "staff@chemo.in", "role": "workshop"}
        database = SimpleNamespace(client=RollbackFakeClient([entries, overrides, revisions]))
        with (
            patch("backend.app.require_role", return_value=(user, None)),
            patch("backend.app.get_database", return_value=database),
            patch("backend.app.get_price_list_imports_collection", return_value=imports),
            patch("backend.app.get_price_list_entries_collection", return_value=entries),
            patch("backend.app.get_price_list_overrides_collection", return_value=overrides),
            patch("backend.app.get_price_list_revisions_collection", return_value=revisions),
        ):
            response = app.test_client().patch(
                f"/price-lists/entries/{entry_id}/price",
                json={"updated_price": 19500},
            )

        self.assertEqual(response.status_code, 503)
        self.assertEqual(overrides.items, [])
        self.assertEqual(revisions.items, [])

    def test_price_revision_returns_503_when_a_collection_is_unavailable(self):
        imports = FakeCollection()
        entries = FakeCollection()
        entry_id = ObjectId()
        imports.items.append({"import_id": "import-1", "active": True, "filename": "prices.csv", "stats": {}})
        entries.items.append({"_id": entry_id, "import_id": "import-1", "price": 18988.0})
        user = {"_id": ObjectId(), "email": "staff@chemo.in", "role": "workshop"}
        with (
            patch("backend.app.require_role", return_value=(user, None)),
            patch("backend.app.get_database", return_value=FakeDatabase()),
            patch("backend.app.get_price_list_imports_collection", return_value=imports),
            patch("backend.app.get_price_list_entries_collection", return_value=entries),
            patch("backend.app.get_price_list_overrides_collection", return_value=None),
            patch("backend.app.get_price_list_revisions_collection", return_value=FakeCollection()),
        ):
            response = app.test_client().patch(
                f"/price-lists/entries/{entry_id}/price",
                json={"updated_price": 19500},
            )

        self.assertEqual(response.status_code, 503)
        self.assertIn("collections are unavailable", response.get_json()["error"])


if __name__ == "__main__":
    unittest.main()

import json
import unittest
from pathlib import Path

from backend.app import (
    CREASING_MATRIX_SIZES,
    CTCP_PLATE_SIZES,
    SPECIALIZED_CHEMICAL_PRODUCTS,
    RUBBER_BLANKET_RULES,
    RUBBER_BLANKET_STOCK_UNIT,
    build_export_rows,
    build_item_payload,
    build_lookup,
    calculate_stock_movement,
    get_item_identity_query,
    process_excel_row,
)


class SpecializedInventoryTests(unittest.TestCase):
    def build(self, payload):
        item, error = build_item_payload(payload)
        self.assertIsNone(error)
        self.assertIsNotNone(item)
        return item

    def test_creasing_matrix_valid_sizes_and_packet_breakdown(self):
        for thickness, sizes in CREASING_MATRIX_SIZES.items():
            item = self.build(
                {
                    "category": "Creasing Matrix",
                    "thickness": thickness,
                    "size": sizes[0],
                    "quantity": 23,
                    "unit": "pkt",
                }
            )
            self.assertEqual(item["quantity"], 23)
            self.assertEqual(item["unit"], "pkt")
            self.assertEqual(item["packaging"]["units_per_box"], 10)

    def test_creasing_matrix_rejects_mismatched_size(self):
        _, error = build_item_payload(
            {
                "category": "Creasing Matrix",
                "thickness": "13",
                "size": "0.3 X 1.0",
                "quantity": 1,
                "unit": "pkt",
            }
        )
        self.assertIn("not valid", error)

    def test_ctcp_sizes_boxes_and_sheets(self):
        for thickness, sizes in CTCP_PLATE_SIZES.items():
            item = self.build(
                {
                    "category": "CTCP Plates",
                    "thickness": thickness,
                    "size": sizes[0],
                    "boxes": 7,
                    "quantity": 7,
                    "total_sheets": 350,
                    "unit": "box",
                }
            )
            self.assertEqual(item["quantity"], 7)
            self.assertEqual(item["packaging"]["sheets_per_box"], 50)

    def test_ctcp_rejects_mismatched_size(self):
        _, error = build_item_payload(
            {
                "category": "CTCP Plates",
                "thickness": "0.20",
                "size": "650 X 550",
                "quantity": 1,
                "unit": "box",
            }
        )
        self.assertIn("not valid", error)

    def test_chemical_container_calculation(self):
        item = self.build(
            {
                "category": "Washing Solutions",
                "product": "Chem R-ol",
                "type": "5L",
                "unit": "L",
                "boxes": 2,
                "loose_units": 3,
                "containers": 11,
                "quantity": 55,
            }
        )
        self.assertEqual(item["quantity"], 55)
        self.assertEqual(item["unit"], "ltr")
        self.assertEqual(item["packaging"]["containers_per_box"], 4)

    def test_all_specialized_chemical_products_are_accepted(self):
        for product_name, config in SPECIALIZED_CHEMICAL_PRODUCTS.items():
            item = self.build(
                {
                    "category": config["category"],
                    "product": product_name,
                    "type": f"{config['pack_size']}{config['unit']}",
                    "unit": config["unit"],
                    "containers_per_box": config["containers_per_box"][0],
                    "boxes": 0,
                    "loose_units": 1,
                    "quantity": config["pack_size"],
                }
            )
            self.assertEqual(item["brand"], product_name)
            self.assertEqual(item["quantity"], config["pack_size"])

    def test_roll_o_clean_preserves_selected_box_configuration(self):
        for containers_per_box in (12, 15, 18):
            item = self.build(
                {
                    "category": "Roller Care Products",
                    "product": "Roll-o-clean",
                    "type": "1kg",
                    "unit": "kg",
                    "containers_per_box": containers_per_box,
                    "boxes": 1,
                    "loose_units": 0,
                    "quantity": containers_per_box,
                }
            )
            self.assertIn(f"/ {containers_per_box} per box", item["type"])

    def test_specialized_stock_movements_are_server_calculated(self):
        chemical = self.build(
            {
                "category": "Washing Solutions",
                "product": "Chem R-ol",
                "type": "5L",
                "unit": "ltr",
                "quantity": 55,
            }
        )
        change, error, _ = calculate_stock_movement(
            chemical, {"movement": {"direction": "out", "boxes": 2, "loose_units": 3}}
        )
        self.assertIsNone(error)
        self.assertEqual(change, -55)

    def test_calibrated_underpacking_paper_normalizes_roll_area(self):
        item = self.build(
            {
                "category": "Calibrated Underpacking Paper",
                "thickness": "0.40",
                "thickness_unit": "mm",
                "width": "1320",
                "width_unit": "mm",
                "length": "100",
                "length_unit": "m",
                "number_of_rolls": 8,
                "unit": "m²",
            }
        )
        self.assertEqual(item["brand"], "__none__")
        self.assertEqual(item["type"], "__none__")
        self.assertEqual(item["quantity"], 1056.0)
        self.assertEqual(item["unit"], "m²")
        self.assertEqual(item["thickness_micron"], 400.0)
        self.assertEqual(item["area_per_roll_sqm"], 132.0)
        self.assertEqual(item["number_of_rolls"], 8)

    def test_calibrated_underpacking_paper_prefixed_category_supports_roll_mode(self):
        item = self.build(
            {
                "category": "05 - Calibrated Underpacking Paper",
                "storage_type": "roll",
                "thickness_micron": 400,
                "width": 1320,
                "width_unit": "mm",
                "length": 100,
                "length_unit": "m",
                "number_of_rolls": 8,
                "unit": "m²",
            }
        )
        self.assertEqual(item["quantity"], 1056.0)
        self.assertEqual(item["unit"], "m²")
        self.assertEqual(item["storage_type"], "roll")
        self.assertEqual(item["thickness_micron"], 400)
        self.assertEqual(item["packaging"]["stock_unit"], "m²")

    def test_calibrated_underpacking_paper_cut_piece_mode_stores_sheets(self):
        item = self.build(
            {
                "category": "05 - Calibrated Underpacking Paper",
                "storage_type": "cut_piece",
                "thickness_micron": 400,
                "width": 1320,
                "width_unit": "mm",
                "length": 795,
                "length_unit": "mm",
                "number_of_sheets": 50,
                "unit": "sheets",
            }
        )
        self.assertEqual(item["quantity"], 50)
        self.assertEqual(item["unit"], "sheets")
        self.assertEqual(item["storage_type"], "cut_piece")
        self.assertEqual(item["number_of_rolls"], None)
        self.assertEqual(item["area_per_sheet_sqm"], 1.0494)
        self.assertEqual(item["packaging"]["stock_unit"], "sheets")

    def test_calibrated_underpacking_paper_cut_piece_validation_and_movement(self):
        base = {
            "category": "05 - Calibrated Underpacking Paper",
            "storage_type": "cut_piece",
            "thickness_micron": 400,
            "width": 1320,
            "width_unit": "mm",
            "length": 795,
            "length_unit": "mm",
            "number_of_sheets": 50,
            "unit": "sheets",
        }
        for invalid in (
            {"storage_type": "unknown"},
            {"thickness_micron": 450},
            {"width": 0},
            {"length_unit": "yard"},
            {"number_of_sheets": 0},
        ):
            with self.subTest(invalid=invalid):
                _, error = build_item_payload({**base, **invalid})
                self.assertIsNotNone(error)

        item = self.build(base)
        change, error, details = calculate_stock_movement(
            item, {"movement": {"direction": "out", "sheets": 10}}
        )
        self.assertIsNone(error)
        self.assertEqual(change, -10)
        self.assertEqual(details["mode"], "sheets")

    def test_calibrated_underpacking_paper_modes_have_distinct_identity(self):
        roll = self.build(
            {
                "category": "05 - Calibrated Underpacking Paper",
                "storage_type": "roll",
                "thickness_micron": 400,
                "width": 1320,
                "width_unit": "mm",
                "length": 100,
                "length_unit": "m",
                "number_of_rolls": 1,
                "unit": "m²",
            }
        )
        cut = self.build(
            {
                "category": "05 - Calibrated Underpacking Paper",
                "storage_type": "cut_piece",
                "thickness_micron": 400,
                "width": 1320,
                "width_unit": "mm",
                "length": 100,
                "length_unit": "m",
                "number_of_sheets": 1,
                "unit": "sheets",
            }
        )
        self.assertNotEqual(get_item_identity_query(roll), get_item_identity_query(cut))

    def test_calibrated_underpacking_paper_cut_piece_excel_recalculates_and_exports(self):
        item, error = process_excel_row(
            {
                "Category": "05 - Calibrated Underpacking Paper",
                "Storage Type": "cut_piece",
                "Thickness (Micron)": 250,
                "Width": 1320,
                "Width Unit": "mm",
                "Length": 795,
                "Length Unit": "mm",
                "No. of Sheets": 50,
                "Quantity": 9999,
                "Unit": "sheets",
            }
        )
        self.assertIsNone(error)
        self.assertEqual(item["quantity"], 50)
        self.assertEqual(item["unit"], "sheets")
        exported = build_export_rows([item])[0]
        self.assertEqual(exported["Storage Type"], "cut_piece")
        self.assertEqual(exported["Thickness (Micron)"], 250)
        self.assertEqual(exported["No. of Sheets"], 50)
        self.assertEqual(exported["Quantity"], 50)

    def test_calibrated_underpacking_paper_accepts_unit_conversions(self):
        item = self.build(
            {
                "category": "Calibrated Underpacking Paper",
                "thickness": 400,
                "thickness_unit": "micron",
                "width": 1.32,
                "width_unit": "m",
                "length": 100000,
                "length_unit": "mm",
                "number_of_rolls": 8,
                "unit": "sqm",
            }
        )
        self.assertEqual(item["quantity"], 1056.0)
        self.assertEqual(item["width_meters"], 1.32)
        self.assertEqual(item["length_meters"], 100.0)
        self.assertEqual(item["thickness_unit"], "micron")

    def test_calibrated_underpacking_paper_rejects_invalid_physical_values(self):
        invalid_payloads = [
            {"width": 0, "width_unit": "mm", "length": 100, "length_unit": "m", "number_of_rolls": 1},
            {"width": 1320, "width_unit": "yard", "length": 100, "length_unit": "m", "number_of_rolls": 1},
            {"width": 1320, "width_unit": "mm", "length": -1, "length_unit": "m", "number_of_rolls": 1},
            {"width": 1320, "width_unit": "mm", "length": 100, "length_unit": "m", "number_of_rolls": 0},
            {"width": 1320, "width_unit": "mm", "length": 100, "length_unit": "m", "number_of_rolls": 1, "thickness": "bad"},
        ]
        for extra in invalid_payloads:
            payload = {
                "category": "Calibrated Underpacking Paper",
                "thickness": 0.40,
                "thickness_unit": "mm",
                "unit": "m²",
                **extra,
            }
            _, error = build_item_payload(payload)
            self.assertIsNotNone(error, extra)

    def test_calibrated_underpacking_paper_roll_movement_uses_area(self):
        item = self.build(
            {
                "category": "Calibrated Underpacking Paper",
                "thickness": 0.40,
                "thickness_unit": "mm",
                "width": 1320,
                "width_unit": "mm",
                "length": 100,
                "length_unit": "m",
                "number_of_rolls": 8,
                "unit": "m²",
            }
        )
        change, error, details = calculate_stock_movement(
            item, {"movement": {"direction": "out", "rolls": 2}}
        )
        self.assertIsNone(error)
        self.assertEqual(change, -264.0)
        self.assertEqual(details["area_per_roll_sqm"], 132.0)
        self.assertEqual(details["rolls"], 2)

    def test_calibrated_underpacking_paper_excel_recalculates_quantity(self):
        rebuilt, error = process_excel_row(
            {
                "Category": "Calibrated Underpacking Paper",
                "Thickness": "400 micron",
                "Width": 1320,
                "Width Unit": "mm",
                "Length": 100,
                "Length Unit": "m",
                "Rolls": 8,
                "Quantity": 1,
                "Unit": "m²",
            }
        )
        self.assertIsNone(error)
        self.assertEqual(rebuilt["quantity"], 1056.0)
        self.assertEqual(rebuilt["thickness_unit"], "micron")

        exported = build_export_rows([rebuilt])[0]
        self.assertEqual(exported["Width Unit"], "mm")
        self.assertEqual(exported["Length Unit"], "m")
        self.assertEqual(exported["Rolls"], 8)
        self.assertEqual(exported["Quantity"], 1056.0)

    def test_excel_export_round_trip_preserves_specialized_metadata(self):
        item = self.build(
            {
                "category": "CTCP Plates",
                "thickness": "0.30",
                "size": "730 X 600",
                "boxes": 4,
                "quantity": 4,
                "total_sheets": 200,
                "unit": "box",
            }
        )
        row = build_export_rows([item])[0]
        rebuilt, error = process_excel_row(row)
        self.assertIsNone(error)
        self.assertEqual(rebuilt["quantity"], 4)
        self.assertEqual(rebuilt["size"], item["size"])

    def test_existing_generic_item_rules_still_work(self):
        item = self.build(
            {
                "category": "Metalback Blankets",
                "brand": "Day",
                "type": "UV",
                "width": "1040",
                "height": "920",
                "thickness": "1.95",
                "quantity": 10,
                "unit": "pcs",
            }
        )
        lookup, error = build_lookup(item)
        self.assertIsNone(error)
        self.assertEqual(lookup["brand"], "Day")

    def test_existing_nine_column_excel_row_still_works(self):
        item, error = process_excel_row(
            {
                "Category": "Metalback Blankets",
                "Brand": "Day",
                "Type": "UV",
                "Batch/Roll No": None,
                "Width": "1040",
                "Length": "920",
                "Thickness": "1.95",
                "Quantity": 10,
                "Unit": "pcs",
            }
        )
        self.assertIsNone(error)
        self.assertEqual(item["quantity"], 10)

    def test_rubber_blanket_uses_actual_width_for_area(self):
        item = self.build(
            {
                "category": "Rubber Blankets",
                "blanket_name": "Magnum SF 1.95",
                "thickness": 1.95,
                "thickness_unit": "mm",
                "nominal_width": 780,
                "actual_width": 790,
                "width_unit": "mm",
                "length": 30,
                "length_unit": "m",
                "roll_no": "3",
                "batch_no": "20222501",
                "number_of_rolls": 1,
                "quantity": 1,
                "unit": "m²",
            }
        )
        self.assertEqual(item["blanket_name"], "Magnum SF 1.95")
        self.assertEqual(item["actual_width"], "790")
        self.assertEqual(item["area_per_roll_sqm"], 23.7)
        self.assertEqual(item["quantity"], 23.7)
        self.assertEqual(item["unit"], "m²")

    def test_rubber_blanket_selectable_thickness_and_unit_conversion(self):
        item = self.build(
            {
                "category": "Rubber Blankets",
                "blanket_name": "Image Sunrise",
                "thickness": 1.70,
                "thickness_unit": "mm",
                "nominal_width": 770,
                "actual_width": 790,
                "length": 28000,
                "length_unit": "mm",
                "number_of_rolls": 2,
                "unit": "sqm",
            }
        )
        self.assertEqual(item["thickness"], "1.70")
        self.assertEqual(item["length_meters"], 28)
        self.assertEqual(item["area_per_roll_sqm"], 22.12)
        self.assertEqual(item["quantity"], 44.24)

    def test_rubber_blanket_alias_print_type_and_ambiguous_width(self):
        item = self.build(
            {
                "category": "Rubber Blankets",
                "blanket_name": "Image Welt Master GR - 1.70 mm",
                "thickness": 1.70,
                "nominal_width": 890,
                "actual_width": 910,
                "length": 28,
                "length_unit": "m",
                "print_type": "Without Print (W/O)",
                "number_of_rolls": 1,
                "unit": "m²",
            }
        )
        self.assertEqual(item["blanket_name"], "Image Web Master GR 1.70")
        self.assertEqual(item["print_type"], "W/O")

    def test_print_master_name_is_canonicalized_from_legacy_alias(self):
        item, error = build_item_payload(
            {
                "category": "Rubber Blankets",
                "blanket_name": "Image Point Master Green",
                "thickness": 1.95,
                "nominal_width": 1030,
                "actual_width": 1070,
                "length": 28,
                "length_unit": "m",
                "print_type": "P",
                "number_of_rolls": 1,
                "unit": RUBBER_BLANKET_STOCK_UNIT,
            }
        )
        self.assertIsNone(error)
        self.assertEqual(item["blanket_name"], "Image Print Master Green")

        _, error = build_item_payload(
            {
                "category": "Rubber Blankets",
                "blanket_name": "Image Web Master GR 1.70",
                "thickness": 1.70,
                "nominal_width": 890,
                "length": 28,
                "length_unit": "m",
                "print_type": "P",
                "number_of_rolls": 1,
                "unit": "m²",
            }
        )
        self.assertIn("actual width is required", error)

    def test_rubber_blanket_rejects_impossible_combinations(self):
        base = {
            "category": "Rubber Blankets",
            "blanket_name": "Magnum SF 1.95",
            "thickness": 1.95,
            "nominal_width": 780,
            "actual_width": 790,
            "length": 30,
            "length_unit": "m",
            "number_of_rolls": 1,
            "unit": "m²",
        }
        invalid_payloads = [
            {**base, "blanket_name": "Unknown Blanket"},
            {**base, "thickness": 1.70},
            {**base, "actual_width": 780},
            {**base, "print_type": "P"},
            {**base, "length": -1},
            {**base, "length_unit": "yard"},
            {**base, "number_of_rolls": 2, "roll_no": "A"},
        ]
        for payload in invalid_payloads:
            with self.subTest(payload=payload):
                _, error = build_item_payload(payload)
                self.assertIsNotNone(error)

    def test_rubber_blanket_identity_keeps_physical_rolls_separate(self):
        base = {
            "category": "Rubber Blankets",
            "blanket_name": "KP UV Black 1.95",
            "thickness": 1.95,
            "nominal_width": 1070,
            "actual_width": 1070,
            "length": 28,
            "length_unit": "m",
            "number_of_rolls": 1,
            "unit": "m²",
        }
        first = self.build({**base, "roll_no": "1"})
        second = self.build({**base, "roll_no": "2"})
        self.assertNotEqual(get_item_identity_query(first), get_item_identity_query(second))

    def test_rubber_blanket_full_and_partial_stock_movements(self):
        item = self.build(
            {
                "category": "Rubber Blankets",
                "blanket_name": "Image Sunrise",
                "thickness": 1.95,
                "nominal_width": 770,
                "actual_width": 790,
                "length": 28,
                "length_unit": "m",
                "number_of_rolls": 2,
                "unit": "m²",
            }
        )
        change, error, details = calculate_stock_movement(
            item, {"movement": {"direction": "out", "rolls": 1}}
        )
        self.assertIsNone(error)
        self.assertEqual(change, -22.12)
        self.assertEqual(details["mode"], "rolls")

        change, error, details = calculate_stock_movement(
            item, {"movement": {"direction": "out", "quantity": 3.5, "unit": "m²"}}
        )
        self.assertIsNone(error)
        self.assertEqual(change, -3.5)
        self.assertEqual(details["mode"], "partial_area")

    def test_rubber_blanket_cut_piece_mode_stores_sheets_and_area_reference(self):
        item = self.build(
            {
                "category": "Rubber Blankets",
                "blanket_name": "Image Sunrise",
                "storage_type": "cut_piece",
                "thickness": 1.70,
                "thickness_unit": "mm",
                "nominal_width": 0.77,
                "actual_width": 0.79,
                "width_unit": "m",
                "length": 795,
                "length_unit": "mm",
                "number_of_sheets": 50,
                "quantity": 9999,
                "unit": "sheets",
            }
        )
        self.assertEqual(item["storage_type"], "cut_piece")
        self.assertEqual(item["quantity"], 50)
        self.assertEqual(item["unit"], "sheets")
        self.assertEqual(item["number_of_rolls"], None)
        self.assertEqual(item["area_per_sheet_sqm"], 0.62805)
        self.assertEqual(item["packaging"]["stock_unit"], "sheets")

        change, error, details = calculate_stock_movement(
            item, {"movement": {"direction": "out", "sheets": 10}}
        )
        self.assertIsNone(error)
        self.assertEqual(change, -10)
        self.assertEqual(details["mode"], "sheets")

    def test_rubber_blanket_storage_modes_have_distinct_identity(self):
        base = {
            "category": "Rubber Blankets",
            "blanket_name": "Image Sunrise",
            "thickness": 1.70,
            "nominal_width": 770,
            "actual_width": 790,
            "length": 28,
            "length_unit": "m",
        }
        roll = self.build({**base, "storage_type": "roll", "number_of_rolls": 1, "unit": "sqm"})
        cut = self.build({**base, "storage_type": "cut_piece", "number_of_sheets": 1, "unit": "sheets"})
        self.assertNotEqual(get_item_identity_query(roll), get_item_identity_query(cut))

    def test_rubber_blanket_cut_piece_excel_recalculates_and_exports(self):
        item, error = process_excel_row(
            {
                "Category": "Rubber Blankets",
                "Blanket Name": "Image Sunrise",
                "Storage Type": "cut_piece",
                "Thickness": 1.70,
                "Thickness Unit": "mm",
                "Nominal Width": 770,
                "Actual Width": 790,
                "Width Unit": "mm",
                "Length": 795,
                "Length Unit": "mm",
                "Number of Sheets": 50,
                "Quantity": 9999,
                "Unit": "sheets",
            }
        )
        self.assertIsNone(error)
        self.assertEqual(item["quantity"], 50)
        exported = build_export_rows([item])[0]
        self.assertEqual(exported["Storage Type"], "cut_piece")
        self.assertEqual(exported["No. of Sheets"], 50)
        self.assertEqual(exported["Area per Sheet"], 0.62805)
        self.assertEqual(exported["Quantity"], 50)

    def test_rubber_blanket_excel_recalculates_quantity_and_exports_fields(self):
        item, error = process_excel_row(
            {
                "Category": "Rubber Blankets",
                "Blanket Name": "SAVA UV Black 1.95",
                "Thickness": 1.95,
                "Thickness Unit": "mm",
                "Nominal Width": 1060,
                "Actual Width": 1070,
                "Width Unit": "mm",
                "Length": 26.75,
                "Length Unit": "m",
                "Roll No": "A2",
                "Batch No": "1306497-A2",
                "Number of Rolls": 1,
                "Quantity": 9999,
                "Unit": "m²",
            }
        )
        self.assertIsNone(error)
        self.assertEqual(item["quantity"], 28.6225)
        row = build_export_rows([item])[0]
        self.assertEqual(row["Blanket Name"], "SAVA UV Black 1.95")
        self.assertEqual(row["Nominal Width"], "1060")
        self.assertEqual(row["Actual Width"], "1070")
        self.assertEqual(row["Roll No"], "A2")
        self.assertEqual(row["Batch No"], "1306497-A2")
        self.assertEqual(row["Area per Roll"], 28.6225)

    def test_rubber_blanket_master_and_frontend_renderer_are_seeded(self):
        self.assertGreaterEqual(len(RUBBER_BLANKET_RULES), 14)
        project_root = Path(__file__).parents[1]
        source = (project_root / "frontend" / "script.js").read_text(encoding="utf-8")
        config = json.loads((project_root / "frontend" / "data" / "inventory-config.json").read_text(encoding="utf-8"))
        configured_names = {entry["name"] for entry in config["rubber_blankets"]}
        self.assertIn('categoryKey === "rubber_blankets"', source)
        self.assertIn("renderRubberBlanketForm", source)
        self.assertIn("inventory-config.json", source)
        self.assertIn("renderM3ZModeFields", source)
        self.assertIn('value="cut_piece"', source)
        self.assertIn("data-rubber-mode-fields", source)
        self.assertIn("rubber_blankets_cut_piece", source)
        self.assertNotIn('name="roll_no"', source)
        self.assertNotIn('name="batch_no"', source)
        self.assertIn("M3Z_THICKNESS_OPTIONS", source)
        self.assertEqual(set(RUBBER_BLANKET_RULES), configured_names)
        self.assertIn("Image Print Master Green", configured_names)
        self.assertNotIn("Image Point Master Green", configured_names)


if __name__ == "__main__":
    unittest.main()

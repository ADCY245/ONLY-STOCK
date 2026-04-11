import os
import re
from datetime import datetime
from io import BytesIO
from pathlib import Path
from pytz import timezone

import pandas as pd
from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
from pymongo.errors import DuplicateKeyError

try:
    from .db import get_database, get_inventory_collection, get_stock_logs_collection
except ImportError:
    from db import get_database, get_inventory_collection, get_stock_logs_collection


app = Flask(__name__)
CORS(app)
FRONTEND_DIR = Path(__file__).resolve().parent.parent / "frontend"

REQUIRED_EXCEL_COLUMNS = [
    "Category",
    "Brand",
    "Type",
    "Width",
    "Height",
    "Thickness",
    "Quantity",
    "Unit",
]
EXCEL_COLUMNS = [
    "Category",
    "Brand",
    "Type",
    "Batch/Roll No",
    "Width",
    "Height",
    "Thickness",
    "Quantity",
    "Unit",
]
DEFAULT_LOW_STOCK_THRESHOLD = 5
THICKNESS_REQUIRED_CATEGORIES = {
    "Rubber Blankets",
    "Metalback Blankets",
    "Underlay Blanket",
    "Calibrated Underpacking Paper",
    "Calibrated Underpacking Film",
    "Creasing Matrix",
    "Cutting Rules",
    "Creasing Rules",
    "Litho Perforation Rules",
}
DIMENSIONAL_CATEGORIES = THICKNESS_REQUIRED_CATEGORIES | {
    "Blanket Barring",
    "Cutting String",
    "Ejection Rubber",
    "Strip Plate",
    "Anti Marking Film",
    "Ink Duct Foil",
    "Productive Foil",
    "Presspahn Sheets",
    "Auto Wash Cloth",
    "ICP Paper",
    "Dampening Hose",
    "Tesamol Tape",
}

NO_BRAND_TYPE_CATEGORIES = {
    "Creasing Matrix",
}

RULE_UNIT_LINKED_CATEGORIES = {
    "Cutting Rules",
    "Creasing Rules",
    "Litho Perforation Rules",
}

CHEMICAL_CATEGORIES = {
    "Washing Solutions",
    "Fountain Solutions",
    "Plate Care Products",
    "Roller Care Products",
    "Blanket Maintenance Products",
}

BLANKET_BATCH_ROLL_CATEGORIES = {
    "Rubber Blankets",
    "Metalback Blankets",
}


def clean_text(value):
    if not isinstance(value, str):
        return None
    cleaned = value.strip()
    return cleaned or None


def parse_integer(value, field_name, allow_negative=False):
    if isinstance(value, bool):
        return None, f"{field_name} must be an integer"

    if isinstance(value, int):
        parsed = value
    elif isinstance(value, float) and value.is_integer():
        parsed = int(value)
    elif isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None, f"{field_name} is required"
        if re.fullmatch(r"-?\d+", stripped) is None:
            return None, f"{field_name} must be an integer"
        parsed = int(stripped)
    else:
        return None, f"{field_name} must be an integer"

    if not allow_negative and parsed < 0:
        return None, f"{field_name} must be a non-negative integer"

    return parsed, None


def parse_number(value, field_name, allow_negative=False):
    if isinstance(value, bool):
        return None, f"{field_name} must be a number"

    if isinstance(value, (int, float)):
        parsed = float(value)
    elif isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None, f"{field_name} is required"
        try:
            parsed = float(stripped)
        except ValueError:
            return None, f"{field_name} must be a number"
    else:
        return None, f"{field_name} must be a number"

    if not allow_negative and parsed < 0:
        return None, f"{field_name} must be a non-negative number"

    return parsed, None


def parse_optional_text(value):
    if value is None:
        return None
    if isinstance(value, str):
        return clean_text(value)
    text = str(value).strip()
    return text or None


def category_requires_brand(category):
    return category not in NO_BRAND_TYPE_CATEGORIES


def category_requires_type(category):
    return category not in NO_BRAND_TYPE_CATEGORIES


def normalize_optional_dimension(value):
    parsed = parse_optional_text(value)
    if parsed is None:
        return None
    compact = re.sub(r"\s+", "", parsed)
    return compact


def normalize_rule_type(value):
    cleaned = clean_text(value) or ""
    lowered = cleaned.lower()
    if lowered in {"packet", "pack", "pkt"}:
        return "pkt"
    if lowered in {"coil", "coils"}:
        return "coil"
    return cleaned


def parse_format_type(value):
    cleaned = clean_text(value)
    if not cleaned:
        return None, None, None, "type is required"

    match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*(ltr|l|kg|g|ml)", cleaned.strip(), flags=re.IGNORECASE)
    if not match:
        return None, None, None, "type must be a format like 1ltr, 5 ltr, 1kg"

    amount = float(match.group(1))
    unit_raw = match.group(2).lower()
    if unit_raw == "l":
        unit_raw = "ltr"
    normalized_type = f"{match.group(1)} {unit_raw}".replace("  ", " ")
    return normalized_type, amount, unit_raw, None


def category_requires_thickness(category):
    return category in THICKNESS_REQUIRED_CATEGORIES


def category_uses_dimensions(category):
    return category in DIMENSIONAL_CATEGORIES


def normalize_dimension(value):
    return normalize_optional_dimension(value)


def build_size_label(width, height):
    if width and height:
        return f"{width} x {height}"
    return width or height or None


def requires_batch_roll_no(category, unit):
    cleaned_unit = clean_text(unit)
    return category in BLANKET_BATCH_ROLL_CATEGORIES and cleaned_unit and cleaned_unit.lower() == "rolls"


def build_item_payload(data):
    category = clean_text(data.get("category"))
    brand = clean_text(data.get("brand"))
    item_type = clean_text(data.get("type"))
    batch_roll_no = parse_optional_text(data.get("batch_roll_no"))
    width = normalize_dimension(data.get("width"))
    height = normalize_dimension(data.get("height"))
    thickness = parse_optional_text(data.get("thickness"))
    unit = clean_text(data.get("unit"))
    if not category:
        return None, "category is required"

    if category in NO_BRAND_TYPE_CATEGORIES:
        brand = "__none__"
        item_type = "__none__"

    requires_brand = category_requires_brand(category)
    requires_type = category_requires_type(category)

    if requires_brand and not brand:
        return None, "brand is required"
    if requires_type and not item_type:
        return None, "type is required"
    if not unit:
        return None, "unit is required"
    if requires_batch_roll_no(category, unit) and not batch_roll_no:
        return None, "batch / roll no. is required for blanket rolls"
    if not requires_batch_roll_no(category, unit):
        batch_roll_no = None

    if category == "Creasing Matrix" and unit.lower() != "pkt":
        return None, "unit must be pkt for this category"

    if category in RULE_UNIT_LINKED_CATEGORIES:
        normalized_type = normalize_rule_type(item_type)
        if normalized_type not in {"coil", "pkt"}:
            return None, "type must be coil or pkt for this category"
        item_type = normalized_type
        if unit.lower() != item_type:
            return None, "unit must match type for this category"

    format_size = None
    format_unit = None
    if category in CHEMICAL_CATEGORIES:
        normalized_type, format_size, format_unit, format_error = parse_format_type(item_type)
        if format_error:
            return None, format_error
        item_type = normalized_type
        if unit.lower() != format_unit:
            return None, "unit must match the type format unit (e.g., ltr or kg)"

    if category in CHEMICAL_CATEGORIES:
        quantity, quantity_error = parse_number(data.get("quantity"), "quantity")
    else:
        quantity, quantity_error = parse_integer(data.get("quantity"), "quantity")

    if category_uses_dimensions(category) and not all([width, height]):
        return None, "width and height are required for this category"

    if category_requires_thickness(category) and not thickness:
        return None, "thickness is required for this category"

    if quantity_error:
        return None, quantity_error

    now = datetime.now(timezone('Asia/Kolkata'))
    payload = {
        "category": category,
        "brand": brand,
        "type": item_type,
        "batch_roll_no": batch_roll_no,
        "width": width,
        "height": height,
        "size": build_size_label(width, height),
        "thickness": thickness,
        "quantity": quantity,
        "unit": unit,
        "created_at": now,
        "updated_at": now,
    }
    if format_size is not None:
        payload["format_size"] = format_size
        payload["format_unit"] = format_unit
    return payload, None


def build_lookup(data):
    category = clean_text(data.get("category"))
    brand = clean_text(data.get("brand"))
    item_type = clean_text(data.get("type"))
    batch_roll_no = parse_optional_text(data.get("batch_roll_no"))
    width = normalize_dimension(data.get("width"))
    height = normalize_dimension(data.get("height"))
    thickness = parse_optional_text(data.get("thickness"))
    unit = clean_text(data.get("unit"))

    if not category:
        return None, "category is required"

    if category in NO_BRAND_TYPE_CATEGORIES:
        brand = "__none__"
        item_type = "__none__"

    if category_requires_brand(category) and not brand:
        return None, "brand is required"
    if category_requires_type(category) and not item_type:
        return None, "type is required"
    if not requires_batch_roll_no(category, unit):
        batch_roll_no = None

    if category_uses_dimensions(category) and not all([width, height]):
        return None, "width and height are required for this category"

    if category_requires_thickness(category) and not thickness:
        return None, "thickness is required for this category"

    return {
        "category": category,
        "brand": brand,
        "type": item_type,
        "batch_roll_no": batch_roll_no,
        "width": width,
        "height": height,
        "thickness": thickness,
    }, None


def build_item_key(lookup):
    return "|".join(
        [
            lookup["category"],
            lookup["brand"],
            lookup["type"],
            lookup.get("batch_roll_no") or "-",
            lookup.get("width") or "-",
            lookup.get("height") or "-",
            lookup.get("thickness") or "-",
        ]
    )


def serialize_item(item):
    return {
        "id": str(item["_id"]),
        "category": item["category"],
        "brand": item["brand"],
        "type": item["type"],
        "batch_roll_no": item.get("batch_roll_no"),
        "width": item.get("width"),
        "height": item.get("height"),
        "size": item.get("size"),
        "thickness": item.get("thickness"),
        "quantity": item["quantity"],
        "unit": item["unit"],
        "created_at": item["created_at"].isoformat() if item.get("created_at") else None,
        "updated_at": item["updated_at"].isoformat() if item.get("updated_at") else None,
    }


def serialize_log(log):
    return {
        "id": str(log["_id"]),
        "item_key": log["item_key"],
        "action": log["action"],
        "category": log["category"],
        "brand": log["brand"],
        "type": log["type"],
        "batch_roll_no": log.get("batch_roll_no"),
        "size": log["size"],
        "width": log.get("width"),
        "height": log.get("height"),
        "thickness": log.get("thickness"),
        "quantity_before": log["quantity_before"],
        "quantity_after": log["quantity_after"],
        "quantity_change": log["quantity_change"],
        "unit": log["unit"],
        "source": log["source"],
        "changed_at": log["changed_at"].isoformat(),
    }


def log_stock_change(item, action, quantity_before, quantity_after, source):
    stock_logs_collection = get_stock_logs_collection()
    stock_logs_collection.insert_one(
        {
            "item_key": build_item_key(item),
            "action": action,
            "category": item["category"],
            "brand": item["brand"],
            "type": item["type"],
            "batch_roll_no": item.get("batch_roll_no"),
            "size": item["size"],
            "width": item.get("width"),
            "height": item.get("height"),
            "thickness": item.get("thickness"),
            "quantity_before": quantity_before,
            "quantity_after": quantity_after,
            "quantity_change": quantity_after - quantity_before,
            "unit": item["unit"],
            "source": source,
            "changed_at": datetime.now(timezone('Asia/Kolkata')),
        }
    )


def create_inventory_query(args):
    query = {}

    category = clean_text(args.get("category"))
    brand = clean_text(args.get("brand"))
    item_type = clean_text(args.get("type"))
    search = clean_text(args.get("search"))
    low_stock = str(args.get("low_stock", "")).lower() in {"1", "true", "yes"}
    thickness = parse_optional_text(args.get("thickness"))

    if category:
        query["category"] = {"$regex": f"^{re.escape(category)}$", "$options": "i"}
    if brand:
        query["brand"] = {"$regex": f"^{re.escape(brand)}$", "$options": "i"}
    if item_type:
        query["type"] = {"$regex": f"^{re.escape(item_type)}$", "$options": "i"}
    if thickness:
        query["thickness"] = {"$regex": f"^{re.escape(thickness)}$", "$options": "i"}

    if search:
        regex = {"$regex": re.escape(search), "$options": "i"}
        query["$or"] = [
            {"category": regex},
            {"brand": regex},
            {"type": regex},
            {"batch_roll_no": regex},
        ]

    if low_stock:
        threshold_value, threshold_error = parse_integer(
            args.get("low_stock_threshold", DEFAULT_LOW_STOCK_THRESHOLD),
            "low_stock_threshold",
        )
        if threshold_error:
            return None, threshold_error
        query["quantity"] = {"$lte": threshold_value}

    return query, None


def process_excel_row(row):
    item, error = build_item_payload(
        {
            "category": row.get("Category"),
            "brand": row.get("Brand"),
            "type": row.get("Type"),
            "batch_roll_no": row.get("Batch/Roll No"),
            "width": row.get("Width"),
            "height": row.get("Height"),
            "thickness": row.get("Thickness"),
            "quantity": row.get("Quantity"),
            "unit": row.get("Unit"),
        }
    )
    return item, error


@app.route("/")
def home():
    return send_from_directory(FRONTEND_DIR, "index.html")


@app.route("/<path:filename>")
def frontend_assets(filename):
    if filename in {"style.css", "script.js"}:
        return send_from_directory(FRONTEND_DIR, filename)
    if filename == "favicon.ico":
        return ("", 204)
    return ("Not Found", 404)


@app.route("/health")
def health():
    try:
        database = get_database()
        database.command("ping")
        return {
            "status": "ok",
            "database": database.name,
        }
    except Exception as exc:
        return (
            {
                "status": "degraded",
                "database": "unavailable",
                "error": str(exc),
            },
            503,
        )


@app.route("/add-item", methods=["POST"])
def add_item():
    inventory_collection = get_inventory_collection()
    data = request.get_json(silent=True) or {}
    item, error = build_item_payload(data)

    if error:
        return jsonify({"error": error}), 400

    try:
        result = inventory_collection.insert_one(item)
    except DuplicateKeyError:
        return jsonify({"error": "Item already exists"}), 409

    item["_id"] = result.inserted_id
    log_stock_change(item, "created", 0, item["quantity"], "manual")
    return jsonify({"message": "Item added", "item": serialize_item(item)}), 201


@app.route("/inventory", methods=["GET"])
def get_inventory():
    inventory_collection = get_inventory_collection()
    query, error = create_inventory_query(request.args)
    if error:
        return jsonify({"error": error}), 400

    items = inventory_collection.find(query).sort(
        [
            ("category", 1),
            ("brand", 1),
            ("type", 1),
            ("batch_roll_no", 1),
            ("width", 1),
            ("height", 1),
            ("thickness", 1),
        ]
    )
    return jsonify([serialize_item(item) for item in items])


@app.route("/stock-logs", methods=["GET"])
def get_stock_logs():
    stock_logs_collection = get_stock_logs_collection()
    limit, error = parse_integer(request.args.get("limit", 50), "limit")
    if error:
        return jsonify({"error": error}), 400

    logs = stock_logs_collection.find().sort("changed_at", -1).limit(limit)
    return jsonify([serialize_log(log) for log in logs])


@app.route("/update-stock", methods=["PUT"])
def update_stock():
    inventory_collection = get_inventory_collection()
    data = request.get_json(silent=True) or {}
    lookup, error = build_lookup(data)
    if error:
        return jsonify({"error": error}), 400

    item = inventory_collection.find_one(lookup)
    if not item:
        return jsonify({"error": "Item not found"}), 404

    if item.get("category") in CHEMICAL_CATEGORIES:
        quantity_change, quantity_error = parse_number(
            data.get("quantity_change"),
            "quantity_change",
            allow_negative=True,
        )
    else:
        quantity_change, quantity_error = parse_integer(
            data.get("quantity_change"),
            "quantity_change",
            allow_negative=True,
        )
    if quantity_error:
        return jsonify({"error": quantity_error}), 400

    new_quantity = item["quantity"] + quantity_change
    if new_quantity < 0:
        return jsonify({"error": "Quantity cannot go below 0"}), 400

    updates = {
        "quantity": new_quantity,
        "updated_at": datetime.now(timezone('Asia/Kolkata')),
    }

    if data.get("thickness") is not None:
        updates["thickness"] = parse_optional_text(data.get("thickness"))

    if data.get("unit") is not None:
        unit = clean_text(data.get("unit"))
        if not unit:
            return jsonify({"error": "unit cannot be empty"}), 400
        updates["unit"] = unit

    updates["size"] = build_size_label(item.get("width"), item.get("height"))

    inventory_collection.update_one(lookup, {"$set": updates})
    updated_item = inventory_collection.find_one(lookup)
    log_item = dict(item)
    log_item["unit"] = updated_item["unit"]
    log_stock_change(
        log_item,
        "updated",
        item["quantity"],
        updated_item["quantity"],
        "manual",
    )

    return jsonify({"message": "Stock updated", "item": serialize_item(updated_item)})


@app.route("/delete-item", methods=["DELETE"])
def delete_item():
    inventory_collection = get_inventory_collection()
    data = request.get_json(silent=True) or {}
    lookup, error = build_lookup(data)
    if error:
        return jsonify({"error": error}), 400

    item = inventory_collection.find_one(lookup)
    if not item:
        return jsonify({"error": "Item not found"}), 404

    inventory_collection.delete_one({"_id": item["_id"]})
    log_stock_change(item, "deleted", item["quantity"], 0, "manual")
    return jsonify({"message": "Item deleted"})


@app.route("/upload-excel", methods=["POST"])
def upload_excel():
    inventory_collection = get_inventory_collection()
    uploaded_file = request.files.get("file")
    if uploaded_file is None or uploaded_file.filename == "":
        return jsonify({"error": "Excel file is required"}), 400

    try:
        dataframe = pd.read_excel(uploaded_file)
    except Exception:
        return jsonify({"error": "Unable to read Excel file"}), 400

    missing_columns = [column for column in REQUIRED_EXCEL_COLUMNS if column not in dataframe.columns]
    if missing_columns:
        return jsonify(
            {
                "error": "Invalid Excel format",
                "missing_columns": missing_columns,
                "expected_columns": EXCEL_COLUMNS,
            }
        ), 400

    selected_columns = [column for column in EXCEL_COLUMNS if column in dataframe.columns]
    records = dataframe[selected_columns].to_dict(orient="records")
    if not records:
        return jsonify({"error": "Excel file is empty"}), 400

    seen_keys = set()
    inserted = 0
    updated = 0

    for index, row in enumerate(records, start=2):
        item, error = process_excel_row(row)
        if error:
            return jsonify({"error": f"Row {index}: {error}"}), 400

        item_key = build_item_key(item)
        if item_key in seen_keys:
            return jsonify({"error": f"Row {index}: duplicate item in Excel file"}), 400
        seen_keys.add(item_key)

        existing_item = inventory_collection.find_one(
            {
                "category": item["category"],
                "brand": item["brand"],
                "type": item["type"],
                "batch_roll_no": item.get("batch_roll_no"),
                "width": item.get("width"),
                "height": item.get("height"),
                "thickness": item.get("thickness"),
            }
        )

        if existing_item:
            inventory_collection.update_one(
                {"_id": existing_item["_id"]},
                {
                    "$set": {
                        "quantity": item["quantity"],
                        "unit": item["unit"],
                        "batch_roll_no": item.get("batch_roll_no"),
                        "size": item["size"],
                        "width": item.get("width"),
                        "height": item.get("height"),
                        "thickness": item.get("thickness"),
                        "updated_at": datetime.now(timezone('Asia/Kolkata')),
                    }
                },
            )
            latest_item = inventory_collection.find_one({"_id": existing_item["_id"]})
            log_stock_change(latest_item, "excel_update", existing_item["quantity"], latest_item["quantity"], "excel")
            updated += 1
        else:
            result = inventory_collection.insert_one(item)
            item["_id"] = result.inserted_id
            log_stock_change(item, "excel_create", 0, item["quantity"], "excel")
            inserted += 1

    return jsonify(
        {
            "message": "Excel processed successfully",
            "inserted": inserted,
            "updated": updated,
            "total_rows": len(records),
        }
    )


@app.route("/export-excel", methods=["GET"])
def export_excel():
    inventory_collection = get_inventory_collection()
    items = list(
        inventory_collection.find().sort(
            [
                ("category", 1),
                ("brand", 1),
                ("type", 1),
                ("batch_roll_no", 1),
                ("width", 1),
                ("height", 1),
                ("thickness", 1),
            ]
        )
    )

    export_rows = [
        {
            "Category": item["category"],
            "Brand": item["brand"],
            "Type": item["type"],
            "Batch/Roll No": item.get("batch_roll_no"),
            "Width": item.get("width"),
            "Height": item.get("height"),
            "Thickness": item.get("thickness"),
            "Quantity": item["quantity"],
            "Unit": item["unit"],
        }
        for item in items
    ]

    dataframe = pd.DataFrame(export_rows, columns=EXCEL_COLUMNS)
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Inventory")

    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="inventory_export.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    debug = os.getenv("FLASK_DEBUG", "").lower() in {"1", "true", "yes"}
    app.run(debug=debug, host="0.0.0.0", port=port)

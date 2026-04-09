import os
import re
from datetime import datetime
from io import BytesIO

import pandas as pd
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from pymongo.errors import DuplicateKeyError

from db import get_database, get_inventory_collection, get_stock_logs_collection


app = Flask(__name__)
CORS(app)

db = get_database()
inventory_collection = get_inventory_collection()
stock_logs_collection = get_stock_logs_collection()

EXCEL_COLUMNS = ["Category", "Brand", "Type", "Size", "Quantity", "Unit"]
DEFAULT_LOW_STOCK_THRESHOLD = 5


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


def build_item_payload(data):
    category = clean_text(data.get("category"))
    brand = clean_text(data.get("brand"))
    item_type = clean_text(data.get("type"))
    size = clean_text(data.get("size"))
    unit = clean_text(data.get("unit"))
    quantity, quantity_error = parse_integer(data.get("quantity"), "quantity")

    if not all([category, brand, item_type, size, unit]):
        return None, "category, brand, type, size, and unit are required"

    if quantity_error:
        return None, quantity_error

    now = datetime.utcnow()
    return {
        "category": category,
        "brand": brand,
        "type": item_type,
        "size": size,
        "quantity": quantity,
        "unit": unit,
        "created_at": now,
        "updated_at": now,
    }, None


def build_lookup(data):
    category = clean_text(data.get("category"))
    brand = clean_text(data.get("brand"))
    item_type = clean_text(data.get("type"))
    size = clean_text(data.get("size"))

    if not all([category, brand, item_type, size]):
        return None, "category, brand, type, and size are required"

    return {
        "category": category,
        "brand": brand,
        "type": item_type,
        "size": size,
    }, None


def build_item_key(lookup):
    return "|".join(
        [lookup["category"], lookup["brand"], lookup["type"], lookup["size"]]
    )


def serialize_item(item):
    return {
        "id": str(item["_id"]),
        "category": item["category"],
        "brand": item["brand"],
        "type": item["type"],
        "size": item["size"],
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
        "size": log["size"],
        "quantity_before": log["quantity_before"],
        "quantity_after": log["quantity_after"],
        "quantity_change": log["quantity_change"],
        "unit": log["unit"],
        "source": log["source"],
        "changed_at": log["changed_at"].isoformat(),
    }


def log_stock_change(item, action, quantity_before, quantity_after, source):
    stock_logs_collection.insert_one(
        {
            "item_key": build_item_key(item),
            "action": action,
            "category": item["category"],
            "brand": item["brand"],
            "type": item["type"],
            "size": item["size"],
            "quantity_before": quantity_before,
            "quantity_after": quantity_after,
            "quantity_change": quantity_after - quantity_before,
            "unit": item["unit"],
            "source": source,
            "changed_at": datetime.utcnow(),
        }
    )


def create_inventory_query(args):
    query = {}

    category = clean_text(args.get("category"))
    brand = clean_text(args.get("brand"))
    item_type = clean_text(args.get("type"))
    search = clean_text(args.get("search"))
    low_stock = str(args.get("low_stock", "")).lower() in {"1", "true", "yes"}

    if category:
        query["category"] = {"$regex": f"^{re.escape(category)}$", "$options": "i"}
    if brand:
        query["brand"] = {"$regex": f"^{re.escape(brand)}$", "$options": "i"}
    if item_type:
        query["type"] = {"$regex": f"^{re.escape(item_type)}$", "$options": "i"}

    if search:
        regex = {"$regex": re.escape(search), "$options": "i"}
        query["$or"] = [
            {"category": regex},
            {"brand": regex},
            {"type": regex},
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
            "size": row.get("Size"),
            "quantity": row.get("Quantity"),
            "unit": row.get("Unit"),
        }
    )
    return item, error


@app.route("/")
def home():
    return "API running"


@app.route("/health")
def health():
    return {
        "status": "ok",
        "database": db.name,
    }


@app.route("/add-item", methods=["POST"])
def add_item():
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
    query, error = create_inventory_query(request.args)
    if error:
        return jsonify({"error": error}), 400

    items = inventory_collection.find(query).sort(
        [("category", 1), ("brand", 1), ("type", 1), ("size", 1)]
    )
    return jsonify([serialize_item(item) for item in items])


@app.route("/stock-logs", methods=["GET"])
def get_stock_logs():
    limit, error = parse_integer(request.args.get("limit", 50), "limit")
    if error:
        return jsonify({"error": error}), 400

    logs = stock_logs_collection.find().sort("changed_at", -1).limit(limit)
    return jsonify([serialize_log(log) for log in logs])


@app.route("/update-stock", methods=["PUT"])
def update_stock():
    data = request.get_json(silent=True) or {}
    lookup, error = build_lookup(data)
    if error:
        return jsonify({"error": error}), 400

    quantity_change, quantity_error = parse_integer(
        data.get("quantity_change"),
        "quantity_change",
        allow_negative=True,
    )
    if quantity_error:
        return jsonify({"error": quantity_error}), 400

    item = inventory_collection.find_one(lookup)
    if not item:
        return jsonify({"error": "Item not found"}), 404

    new_quantity = item["quantity"] + quantity_change
    if new_quantity < 0:
        return jsonify({"error": "Quantity cannot go below 0"}), 400

    updates = {
        "quantity": new_quantity,
        "updated_at": datetime.utcnow(),
    }

    if data.get("unit") is not None:
        unit = clean_text(data.get("unit"))
        if not unit:
            return jsonify({"error": "unit cannot be empty"}), 400
        updates["unit"] = unit

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
    uploaded_file = request.files.get("file")
    if uploaded_file is None or uploaded_file.filename == "":
        return jsonify({"error": "Excel file is required"}), 400

    try:
        dataframe = pd.read_excel(uploaded_file)
    except Exception:
        return jsonify({"error": "Unable to read Excel file"}), 400

    missing_columns = [column for column in EXCEL_COLUMNS if column not in dataframe.columns]
    if missing_columns:
        return jsonify(
            {
                "error": "Invalid Excel format",
                "missing_columns": missing_columns,
                "expected_columns": EXCEL_COLUMNS,
            }
        ), 400

    records = dataframe[EXCEL_COLUMNS].to_dict(orient="records")
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
                "size": item["size"],
            }
        )

        if existing_item:
            inventory_collection.update_one(
                {"_id": existing_item["_id"]},
                {
                    "$set": {
                        "quantity": item["quantity"],
                        "unit": item["unit"],
                        "updated_at": datetime.utcnow(),
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
    items = list(
        inventory_collection.find().sort(
            [("category", 1), ("brand", 1), ("type", 1), ("size", 1)]
        )
    )

    export_rows = [
        {
            "Category": item["category"],
            "Brand": item["brand"],
            "Type": item["type"],
            "Size": item["size"],
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
    app.run(debug=True, host="0.0.0.0", port=port)

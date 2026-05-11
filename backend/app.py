import os
import re
from datetime import datetime, timezone as datetime_timezone
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
from bson import ObjectId
from flask import Flask, jsonify, request, send_file, send_from_directory, session
from flask_cors import CORS
from pymongo.errors import DuplicateKeyError
from werkzeug.security import check_password_hash, generate_password_hash

try:
    from .db import get_database, get_inventory_collection, get_stock_logs_collection, get_users_collection
except ImportError:
    from db import get_database, get_inventory_collection, get_stock_logs_collection, get_users_collection


app = Flask(__name__)
CORS(app, supports_credentials=True)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "only-stock-dev-secret")
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
FRONTEND_DIR = Path(__file__).resolve().parent.parent / "frontend"
ALLOWED_SIGNUP_DOMAIN = "@chemo.in"
SUPERADMIN_EMAIL = "athulnair3096@gmail.com"
SUPERADMIN_PASSWORD = "ADMIN123"
ALLOWED_ROLES = {"user", "admin", "workshop"}
READ_ONLY_ROLES = {"user"}
WRITE_ROLES = {"admin", "workshop"}
LOG_ACCESS_ROLES = {"admin"}

REQUIRED_EXCEL_COLUMNS = [
    "Category",
    "Brand",
    "Type",
    "Width",
    "Length",
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
    "Length",
    "Thickness",
    "Quantity",
    "Unit",
]
DEFAULT_LOW_STOCK_THRESHOLD = 5
IST_TIMEZONE = ZoneInfo("Asia/Kolkata")
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


def now_ist():
    return datetime.now(IST_TIMEZONE)


def serialize_datetime_ist(value):
    if not value:
        return None
    if value.tzinfo is None:
        value = value.replace(tzinfo=datetime_timezone.utc)
    return value.astimezone(IST_TIMEZONE).isoformat()


def clean_text(value):
    if not isinstance(value, str):
        return None
    cleaned = value.strip()
    return cleaned or None


def normalize_email(value):
    cleaned = clean_text(value)
    return cleaned.lower() if cleaned else None


def serialize_user(user):
    return {
        "id": str(user["_id"]),
        "email": user["email"],
        "role": user["role"],
        "created_at": serialize_datetime_ist(user.get("created_at")),
        "updated_at": serialize_datetime_ist(user.get("updated_at")),
    }


def get_current_user():
    user_id = session.get("user_id")
    if not user_id:
        return None
    try:
        return get_users_collection().find_one({"_id": ObjectId(user_id)})
    except Exception:
        session.clear()
        return None


def require_auth():
    user = get_current_user()
    if not user:
        return None, (jsonify({"error": "Authentication required"}), 401)
    return user, None


def require_role(*roles):
    user, error_response = require_auth()
    if error_response:
        return None, error_response
    if user["role"] not in roles:
        return None, (jsonify({"error": "You do not have permission for this action"}), 403)
    return user, None


def ensure_superadmin():
    """Ensure superadmin user exists with the configured password."""
    users_collection = get_users_collection()
    user = users_collection.find_one({"email": SUPERADMIN_EMAIL})

    now = now_ist()
    if user:
        # Update password to ensure it matches SUPERADMIN_PASSWORD
        users_collection.update_one(
            {"_id": user["_id"]},
            {"$set": {
                "password_hash": generate_password_hash(SUPERADMIN_PASSWORD),
                "role": "admin",
                "updated_at": now,
            }}
        )
        return {"created": False, "updated": True, "email": SUPERADMIN_EMAIL}
    else:
        # Create new superadmin
        user = {
            "email": SUPERADMIN_EMAIL,
            "password_hash": generate_password_hash(SUPERADMIN_PASSWORD),
            "role": "admin",
            "created_at": now,
            "updated_at": now,
        }
        result = users_collection.insert_one(user)
        return {"created": True, "updated": False, "email": SUPERADMIN_EMAIL, "id": str(result.inserted_id)}


@app.route("/init-superadmin", methods=["POST"])
def init_superadmin():
    """Endpoint to initialize/create the superadmin user."""
    try:
        result = ensure_superadmin()
        return jsonify({"message": "Superadmin ensured", "result": result})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


def validate_signup_email(email):
    if email == SUPERADMIN_EMAIL:
        return None
    if not email.endswith(ALLOWED_SIGNUP_DOMAIN):
        return "Signup is not permitted for this email"
    return None


def validate_password(password):
    if not isinstance(password, str) or len(password.strip()) < 6:
        return "Password must be at least 6 characters"
    return None


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


def parse_dimension_number(value):
    text = str(value or "").strip()
    match = re.search(r"\d+(?:\.\d+)?", text)
    if not match:
        return None
    return float(match.group(0))


def calculate_roll_area_sqm(width, length):
    width_mm = parse_dimension_number(width)
    length_mtr = parse_dimension_number(length)
    if width_mm is None or length_mtr is None:
        return None
    return round((width_mm / 1000) * length_mtr, 4)


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


def build_size_label(width, length):
    if width and length:
        return f"{width} x {length}"
    return width or length or None


def requires_batch_roll_no(category, unit):
    cleaned_unit = clean_text(unit)
    return category in BLANKET_BATCH_ROLL_CATEGORIES and cleaned_unit and cleaned_unit.lower() == "rolls"


def is_roll_unit(unit):
    cleaned_unit = clean_text(unit)
    return cleaned_unit is not None and cleaned_unit.lower() == "rolls"


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

    if is_roll_unit(unit):
        quantity = 0
        quantity_error = None
    elif category in CHEMICAL_CATEGORIES:
        quantity, quantity_error = parse_number(data.get("quantity"), "quantity")
    else:
        quantity, quantity_error = parse_integer(data.get("quantity"), "quantity")

    if category_uses_dimensions(category) and not all([width, height]):
        return None, "width and length are required for this category"

    if category_requires_thickness(category) and not thickness:
        return None, "thickness is required for this category"

    if quantity_error:
        return None, quantity_error
    if is_roll_unit(unit):
        roll_area = calculate_roll_area_sqm(width, height)
        if roll_area is None:
            return None, "width and length must be numeric for roll sq.m calculation"
        quantity = roll_area

    now = now_ist()
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
        return None, "width and length are required for this category"

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
        "created_at": serialize_datetime_ist(item.get("created_at")),
        "updated_at": serialize_datetime_ist(item.get("updated_at")),
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
        "reason": log.get("reason"),
        "details": log.get("details") or {},
        "changed_at": serialize_datetime_ist(log.get("changed_at")),
    }


def log_stock_change(item, action, quantity_before, quantity_after, source, reason=None, details=None):
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
            "reason": parse_optional_text(reason),
            "details": details or {},
            "changed_at": now_ist(),
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
            "height": row.get("Length", row.get("Height")),
            "thickness": row.get("Thickness"),
            "quantity": row.get("Quantity"),
            "unit": row.get("Unit"),
        }
    )
    return item, error


def get_request_reason(data, default_reason=None):
    reason = parse_optional_text(data.get("reason"))
    return reason or default_reason


def get_excel_reason():
    return parse_optional_text(request.form.get("reason")) or "Excel upload"


def get_item_identity_query(item):
    return {
        "category": item["category"],
        "brand": item["brand"],
        "type": item["type"],
        "batch_roll_no": item.get("batch_roll_no"),
        "width": item.get("width"),
        "height": item.get("height"),
        "thickness": item.get("thickness"),
    }


def get_inventory_sort():
    return [
        ("category", 1),
        ("brand", 1),
        ("type", 1),
        ("batch_roll_no", 1),
        ("width", 1),
        ("height", 1),
        ("thickness", 1),
    ]


def build_export_rows(items):
    return [
        {
            "Category": item["category"],
            "Brand": item["brand"],
            "Type": item["type"],
            "Batch/Roll No": item.get("batch_roll_no"),
            "Width": item.get("width"),
            "Length": item.get("height"),
            "Thickness": item.get("thickness"),
            "Quantity": item["quantity"],
            "Unit": item["unit"],
        }
        for item in items
    ]


def send_excel(dataframe, download_name, sheet_name):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name=sheet_name)

    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def get_changed_fields(existing_item, item):
    comparable_fields = [
        "quantity",
        "unit",
        "batch_roll_no",
        "size",
        "width",
        "height",
        "thickness",
    ]
    changes = {}
    for field in comparable_fields:
        before = existing_item.get(field)
        after = item.get(field)
        if before != after:
            changes[field] = {"before": before, "after": after}
    return changes


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


@app.route("/auth/me", methods=["GET"])
def auth_me():
    user, error_response = require_auth()
    if error_response:
        return error_response
    return jsonify({"user": serialize_user(user)})


@app.route("/auth/signup", methods=["POST"])
def auth_signup():
    users_collection = get_users_collection()
    data = request.get_json(silent=True) or {}
    email = normalize_email(data.get("email"))
    password = data.get("password")
    if not email:
        return jsonify({"error": "Email is required"}), 400

    email_error = validate_signup_email(email)
    if email_error:
        return jsonify({"error": email_error}), 400

    password_error = validate_password(password)
    if password_error:
        return jsonify({"error": password_error}), 400

    role = "user"
    if email == SUPERADMIN_EMAIL:
        role = "admin"

    now = now_ist()
    user = {
        "email": email,
        "password_hash": generate_password_hash(password.strip()),
        "role": role,
        "created_at": now,
        "updated_at": now,
    }

    try:
        result = users_collection.insert_one(user)
    except DuplicateKeyError:
        return jsonify({"error": "An account with this email already exists"}), 409

    user["_id"] = result.inserted_id
    session["user_id"] = str(result.inserted_id)
    return jsonify({"message": "Signup successful", "user": serialize_user(user)}), 201


@app.route("/auth/login", methods=["POST"])
def auth_login():
    data = request.get_json(silent=True) or {}
    email = normalize_email(data.get("email"))
    password = data.get("password")

    if not email or not isinstance(password, str):
        return jsonify({"error": "Email and password are required"}), 400

    if email != SUPERADMIN_EMAIL and not email.endswith(ALLOWED_SIGNUP_DOMAIN):
        return jsonify({"error": "Not permitted to login"}), 403

    user = get_users_collection().find_one({"email": email})
    if not user or not check_password_hash(user["password_hash"], password):
        return jsonify({"error": "Invalid email or password"}), 401

    session["user_id"] = str(user["_id"])
    return jsonify({"message": "Login successful", "user": serialize_user(user)})


@app.route("/auth/logout", methods=["POST"])
def auth_logout():
    session.clear()
    return jsonify({"message": "Logged out"})


@app.route("/auth/forgot-password", methods=["POST"])
def auth_forgot_password():
    data = request.get_json(silent=True) or {}
    email = normalize_email(data.get("email"))
    new_password = data.get("new_password")

    if not email:
        return jsonify({"error": "Email is required"}), 400

    password_error = validate_password(new_password)
    if password_error:
        return jsonify({"error": password_error}), 400

    user = get_users_collection().find_one({"email": email})
    if not user:
        return jsonify({"error": "Account not found"}), 404

    get_users_collection().update_one(
        {"_id": user["_id"]},
        {
            "$set": {
                "password_hash": generate_password_hash(new_password.strip()),
                "updated_at": now_ist(),
            }
        },
    )
    session["user_id"] = str(user["_id"])
    updated_user = get_users_collection().find_one({"_id": user["_id"]})
    return jsonify({"message": "Password reset successful", "user": serialize_user(updated_user)})


@app.route("/add-item", methods=["POST"])
def add_item():
    _, error_response = require_role(*WRITE_ROLES)
    if error_response:
        return error_response
    inventory_collection = get_inventory_collection()
    data = request.get_json(silent=True) or {}
    item, error = build_item_payload(data)
    reason = get_request_reason(data, "Manual item creation")

    if error:
        return jsonify({"error": error}), 400

    try:
        result = inventory_collection.insert_one(item)
    except DuplicateKeyError:
        return jsonify({"error": "Item already exists"}), 409

    item["_id"] = result.inserted_id
    log_stock_change(item, "created", 0, item["quantity"], "manual", reason)
    return jsonify({"message": "Item added", "item": serialize_item(item)}), 201


@app.route("/inventory", methods=["GET"])
def get_inventory():
    _, error_response = require_role("user", "workshop", "admin")
    if error_response:
        return error_response
    inventory_collection = get_inventory_collection()
    query, error = create_inventory_query(request.args)
    if error:
        return jsonify({"error": error}), 400

    items = inventory_collection.find(query).sort(get_inventory_sort())
    return jsonify([serialize_item(item) for item in items])


@app.route("/stock-logs", methods=["GET"])
def get_stock_logs():
    _, error_response = require_role(*LOG_ACCESS_ROLES)
    if error_response:
        return error_response
    stock_logs_collection = get_stock_logs_collection()
    limit, error = parse_integer(request.args.get("limit", 50), "limit")
    if error:
        return jsonify({"error": error}), 400

    logs = stock_logs_collection.find().sort("changed_at", -1).limit(limit)
    return jsonify([serialize_log(log) for log in logs])


@app.route("/update-stock", methods=["PUT"])
def update_stock():
    _, error_response = require_role(*WRITE_ROLES)
    if error_response:
        return error_response
    inventory_collection = get_inventory_collection()
    data = request.get_json(silent=True) or {}
    lookup, error = build_lookup(data)
    reason = get_request_reason(data, "Manual stock movement")
    if error:
        return jsonify({"error": error}), 400

    item = inventory_collection.find_one(lookup)
    if not item:
        return jsonify({"error": "Item not found"}), 404

    if item.get("category") in CHEMICAL_CATEGORIES or is_roll_unit(item.get("unit")):
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
        "updated_at": now_ist(),
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
        reason,
        {
            "movement": data.get("quantity_change"),
            "quantity_before": item["quantity"],
            "quantity_after": updated_item["quantity"],
        },
    )

    return jsonify({"message": "Stock updated", "item": serialize_item(updated_item)})


@app.route("/delete-item", methods=["DELETE"])
def delete_item():
    _, error_response = require_role(*WRITE_ROLES)
    if error_response:
        return error_response
    inventory_collection = get_inventory_collection()
    data = request.get_json(silent=True) or {}
    lookup, error = build_lookup(data)
    reason = get_request_reason(data, "Manual item deletion")
    if error:
        return jsonify({"error": error}), 400

    item = inventory_collection.find_one(lookup)
    if not item:
        return jsonify({"error": "Item not found"}), 404

    inventory_collection.delete_one({"_id": item["_id"]})
    log_stock_change(item, "deleted", item["quantity"], 0, "manual", reason)
    return jsonify({"message": "Item deleted"})


@app.route("/upload-excel", methods=["POST"])
def upload_excel():
    _, error_response = require_role(*WRITE_ROLES)
    if error_response:
        return error_response
    inventory_collection = get_inventory_collection()
    uploaded_file = request.files.get("file")
    upload_mode = (request.form.get("mode") or "import").strip().lower()
    reason = get_excel_reason()
    if upload_mode not in {"import", "update"}:
        return jsonify({"error": "mode must be import or update"}), 400

    if uploaded_file is None or uploaded_file.filename == "":
        return jsonify({"error": "Excel file is required"}), 400

    try:
        dataframe = pd.read_excel(uploaded_file)
    except Exception:
        return jsonify({"error": "Unable to read Excel file"}), 400

    missing_columns = []
    for column in REQUIRED_EXCEL_COLUMNS:
        if column == "Length":
            if "Length" not in dataframe.columns and "Height" not in dataframe.columns:
                missing_columns.append(column)
        elif column not in dataframe.columns:
            missing_columns.append(column)
    if missing_columns:
        return jsonify(
            {
                "error": "Invalid Excel format",
                "missing_columns": missing_columns,
                "expected_columns": EXCEL_COLUMNS,
            }
        ), 400

    selected_columns = [column for column in EXCEL_COLUMNS if column in dataframe.columns]
    if "Length" not in selected_columns and "Height" in dataframe.columns:
        selected_columns.append("Height")
    records = dataframe[selected_columns].to_dict(orient="records")
    if not records:
        return jsonify({"error": "Excel file is empty"}), 400

    seen_keys = set()
    inserted = 0
    updated = 0
    unchanged = 0

    for index, row in enumerate(records, start=2):
        item, error = process_excel_row(row)
        if error:
            return jsonify({"error": f"Row {index}: {error}"}), 400

        item_key = build_item_key(item)
        if item_key in seen_keys:
            return jsonify({"error": f"Row {index}: duplicate item in Excel file"}), 400
        seen_keys.add(item_key)

        existing_item = inventory_collection.find_one(get_item_identity_query(item))

        if existing_item:
            changes = get_changed_fields(existing_item, item)
            if not changes:
                unchanged += 1
                continue

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
                        "updated_at": now_ist(),
                    }
                },
            )
            latest_item = inventory_collection.find_one({"_id": existing_item["_id"]})
            log_stock_change(
                latest_item,
                "excel_update",
                existing_item["quantity"],
                latest_item["quantity"],
                "excel",
                reason,
                {"mode": upload_mode, "changes": changes},
            )
            updated += 1
        else:
            result = inventory_collection.insert_one(item)
            item["_id"] = result.inserted_id
            log_stock_change(item, "excel_create", 0, item["quantity"], "excel", reason, {"mode": upload_mode})
            inserted += 1

    deleted = 0
    if upload_mode == "update":
        for existing_item in inventory_collection.find():
            if build_item_key(existing_item) in seen_keys:
                continue
            inventory_collection.delete_one({"_id": existing_item["_id"]})
            log_stock_change(
                existing_item,
                "excel_delete",
                existing_item["quantity"],
                0,
                "excel",
                reason,
                {"mode": upload_mode, "delete_source": "missing_from_update_sheet"},
            )
            deleted += 1

    return jsonify(
        {
            "message": "Excel processed successfully",
            "inserted": inserted,
            "updated": updated,
            "deleted": deleted,
            "unchanged": unchanged,
            "total_rows": len(records),
            "mode": upload_mode,
        }
    )


@app.route("/download-import-template", methods=["GET"])
def download_import_template():
    _, error_response = require_role(*WRITE_ROLES)
    if error_response:
        return error_response
    dataframe = pd.DataFrame(columns=EXCEL_COLUMNS)
    return send_excel(dataframe, "import_items_template.xlsx", "Import Items")


@app.route("/export-update-excel", methods=["GET"])
def export_update_excel():
    _, error_response = require_role(*WRITE_ROLES)
    if error_response:
        return error_response
    inventory_collection = get_inventory_collection()
    items = list(inventory_collection.find().sort(get_inventory_sort()))
    export_rows = build_export_rows(items)
    dataframe = pd.DataFrame(export_rows, columns=EXCEL_COLUMNS)
    return send_excel(dataframe, "update_items_current_stock.xlsx", "Update Items")


@app.route("/export-excel", methods=["GET"])
def export_excel():
    return export_update_excel()


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    debug = os.getenv("FLASK_DEBUG", "").lower() in {"1", "true", "yes"}
    app.run(debug=debug, host="0.0.0.0", port=port)

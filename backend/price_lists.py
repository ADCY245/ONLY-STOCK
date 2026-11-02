import csv
import io
import re
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Side, Border
from openpyxl.utils import get_column_letter


PRICE_LIST_FIELDS = {
    "price_list_id": "PriceList ID",
    "customer_name": "PriceList Name",
    "item_name": "Item Name",
    "price": "PriceList Rate",
}

FIELD_ALIASES = {
    "price_list_id": ("pricelistid", "pricelistno", "pricelistnumber", "listid", "priceid", "reference", "referenceid", "id"),
    "customer_name": ("pricelistname", "customer", "customername", "client", "account", "accountname", "partyname"),
    "item_name": ("itemname", "product", "productname", "description", "itemdescription", "productdescription"),
    "price": ("pricelistrate", "rate", "price", "customerprice", "sellingprice", "unitprice", "amount"),
}

PRICE_LIST_HEADERS = [
    "Brand", "Product Type", "Size / Code", "Original Product Description",
    "Original Price", "Updated Pricing", "PriceList ID",
]
PRICE_LIST_HEADERS_WITH_CUSTOMER = ["Customer", *PRICE_LIST_HEADERS]

# These defaults are intentionally small and extendable. The parser still uses
# the first token as a fallback brand and keeps the original description intact.
DEFAULT_PRODUCT_PARSER_CONFIG = {
    "dimension_pattern": (
        r"(?ix)\b\d+(?:\.\d+)?\s*(?:mm|mtr|meter|metre|m|cm|inch|in|kg|kgs|l|ltr|litre|g|gsm|micron|um|µm)"
        r"(?:\s*[x×]\s*\d+(?:\.\d+)?\s*(?:mm|mtr|meter|metre|m|cm|inch|in|kg|kgs|l|ltr|litre|g|gsm|micron|um|µm))*"
    ),
    "code_pattern": r"\b\d{4,}\b",
    "brand_aliases": {},
}


def normalize_header(value):
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def detect_column_mapping(headers, overrides=None):
    headers = [str(header).strip() for header in headers if str(header).strip()]
    normalized = {normalize_header(header): header for header in headers}
    overrides = overrides or {}
    mapping = {}
    for field in PRICE_LIST_FIELDS:
        requested = overrides.get(field)
        if requested in headers:
            mapping[field] = requested
            continue
        requested_normalized = normalize_header(requested)
        if requested_normalized in normalized:
            mapping[field] = normalized[requested_normalized]
            continue
        aliases = FIELD_ALIASES[field]
        mapping[field] = next((normalized[alias] for alias in aliases if alias in normalized), None)

    # Positional fallback supports files with generic or missing headers.
    positions = {field: index for index, field in enumerate(PRICE_LIST_FIELDS)}
    has_named_match = any(mapping.get(field) for field in ("customer_name", "item_name", "price"))
    for field, index in positions.items():
        if field == "price_list_id" and (len(headers) < 4 or has_named_match):
            continue
        if mapping[field] is None and index < len(headers):
            mapping[field] = headers[index]
    return mapping


def _raw_text(value):
    if value is None:
        return ""
    try:
        if bool(pd.isna(value)):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value)


def _clean_customer(value):
    return re.sub(r"\s+", " ", _raw_text(value)).strip()


def parse_price_value(value):
    text = _raw_text(value).strip()
    if not text:
        return None, "Price is blank"
    cleaned = re.sub(r"[₹$€£,\s]", "", text)
    try:
        number = Decimal(cleaned)
    except InvalidOperation:
        return None, "Price must be numeric"
    if not number.is_finite() or number < 0:
        return None, "Price must be a non-negative number"
    return float(number), None


def _parse_price(value):
    return parse_price_value(value)


def _product_parser_config(config=None):
    merged = dict(DEFAULT_PRODUCT_PARSER_CONFIG)
    if config:
        merged.update({key: value for key, value in config.items() if value is not None})
    return merged


def parse_product_description(description, config=None):
    original = _raw_text(description)
    visible = original.strip()
    if not visible:
        return {"brand": None, "product_type": None, "size_code": None, "original_item_description": original}

    parser_config = _product_parser_config(config)
    aliases = {normalize_header(key): str(value).strip() for key, value in (parser_config.get("brand_aliases") or {}).items()}
    tokens = visible.split()
    first_token = tokens[0].strip("-–—|,;:")
    brand = aliases.get(normalize_header(first_token), first_token) or None

    remainder = visible[len(tokens[0]):].strip()
    size_match = re.search(parser_config["dimension_pattern"], remainder)
    size_code = None
    if size_match:
        size_code = size_match.group(0).strip()
        product_type = remainder[:size_match.start()].strip(" -|,;:") or None
    else:
        code_match = re.search(parser_config["code_pattern"], remainder)
        if code_match:
            size_code = code_match.group(0)
            product_type = remainder[:code_match.start()].strip(" -|,;:") or None
        else:
            product_type = remainder.strip(" -|,;:") or None

    return {
        "brand": brand,
        "product_type": product_type,
        "size_code": size_code,
        "original_item_description": original,
    }


def _read_dataframe(file_bytes, filename):
    suffix = str(filename or "").lower().rsplit(".", 1)[-1]
    if suffix == "csv":
        last_error = None
        for encoding in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                return pd.read_csv(BytesIO(file_bytes), dtype=object, keep_default_na=False, encoding=encoding)
            except UnicodeDecodeError as error:
                last_error = error
        raise ValueError(f"Unable to decode CSV file: {last_error}")
    if suffix in {"xlsx", "xls"}:
        return pd.read_excel(BytesIO(file_bytes), dtype=object)
    raise ValueError("Supported file formats are .csv, .xlsx, and .xls")


def parse_price_list_file(file_storage, mapping_overrides=None, parser_config=None):
    filename = getattr(file_storage, "filename", "") or "price-list"
    file_storage.stream.seek(0)
    file_bytes = file_storage.stream.read()
    if not file_bytes:
        raise ValueError("The uploaded file is empty")
    dataframe = _read_dataframe(file_bytes, filename)
    headers = [str(column).strip() for column in dataframe.columns]
    mapping = detect_column_mapping(headers, mapping_overrides)
    missing = [PRICE_LIST_FIELDS[field] for field in ("customer_name", "item_name", "price") if not mapping.get(field)]
    if missing:
        raise ValueError("Unable to map required columns: " + ", ".join(missing))

    valid_records = []
    invalid_records = []
    seen = set()
    duplicate_records = []
    observed_customers = set()
    observed_products = set()
    for offset, row in dataframe.iterrows():
        row_number = int(offset) + 2
        customer = _clean_customer(row.get(mapping["customer_name"]))
        original_item = _raw_text(row.get(mapping["item_name"]))
        if customer:
            observed_customers.add(customer)
        if original_item.strip():
            observed_products.add(original_item)
        price, price_error = _parse_price(row.get(mapping["price"]))
        reasons = []
        if not customer:
            reasons.append("Customer name is blank")
        if not original_item.strip():
            reasons.append("Item description is blank")
        if price_error:
            reasons.append(price_error)
        if reasons:
            invalid_records.append({
                "row": row_number,
                "customer": customer,
                "item": original_item,
                "price": _raw_text(row.get(mapping["price"])),
                "reasons": reasons,
            })
            continue

        parsed = parse_product_description(original_item, parser_config)
        record = {
            "row_number": row_number,
            "price_list_id": _raw_text(row.get(mapping["price_list_id"])).strip() or None,
            "customer_name": customer,
            "item_name": original_item,
            "price": price,
            **parsed,
        }
        duplicate_key = (record["price_list_id"], customer, original_item, price)
        if duplicate_key in seen:
            duplicate_records.append({"row": row_number, "customer": customer, "item": original_item, "price": price})
            continue
        seen.add(duplicate_key)
        valid_records.append(record)

    stats = {
        "total_records": int(len(dataframe.index)),
        "unique_customers": len(observed_customers),
        "unique_products": len(observed_products),
        "invalid_records": len(invalid_records),
        "duplicate_records": len(duplicate_records),
        "successful_import_count": len(valid_records),
    }
    return {
        "filename": filename,
        "headers": headers,
        "mapping": mapping,
        "records": valid_records,
        "invalid_records": invalid_records,
        "duplicate_records": duplicate_records,
        "stats": stats,
    }


def _safe_sheet_name(name, used_names):
    cleaned = re.sub(r"[\\/*?:\[\]]", "_", str(name or "Customer")).strip(" '") or "Customer"
    cleaned = cleaned[:31]
    base = cleaned
    index = 2
    while cleaned in used_names:
        suffix = f"_{index}"
        cleaned = f"{base[:31 - len(suffix)]}{suffix}"
        index += 1
    used_names.add(cleaned)
    return cleaned


def build_price_list_workbook(records, include_customer=False, title="Price List"):
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)
    grouped = defaultdict(list)
    if include_customer:
        grouped[title] = list(records)
    else:
        for record in records:
            grouped[record["customer_name"]].append(record)
    if not grouped:
        grouped[title] = []

    used_names = set()
    header_fill = PatternFill("solid", fgColor="113943")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(bottom=Side(style="thin", color="D8C7AF"))
    headers = PRICE_LIST_HEADERS_WITH_CUSTOMER if include_customer else PRICE_LIST_HEADERS
    for sheet_title, sheet_records in grouped.items():
        sheet = workbook.create_sheet(_safe_sheet_name(sheet_title, used_names))
        sheet.append(headers)
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="left")
            cell.border = border
        for record in sheet_records:
            original_price = record.get("original_price", record.get("price"))
            updated_price = record.get("updated_price", original_price)
            values = [
                record["brand"], record["product_type"], record["size_code"],
                record["original_item_description"], original_price, updated_price,
                record.get("price_list_id"),
            ]
            if include_customer:
                values.insert(0, record["customer_name"])
            sheet.append(values)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        first_price_column = 6 if include_customer else 5
        for column in (first_price_column, first_price_column + 1):
            for cell in sheet.iter_cols(min_col=column, max_col=column, min_row=2):
                for price_cell in cell:
                    price_cell.number_format = "#,##0.00"
        widths = [22, 26, 24, 62, 16, 16, 22] if not include_customer else [30, 22, 26, 24, 62, 16, 16, 22]
        for index, width in enumerate(widths, start=1):
            sheet.column_dimensions[get_column_letter(index)].width = width
        for row in sheet.iter_rows():
            for cell in row:
                wrap_columns = {4} if not include_customer else {5}
                cell.alignment = Alignment(vertical="top", wrap_text=cell.column in wrap_columns)
    return workbook


def build_price_revision_history_workbook(records):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Price Revision History"
    headers = [
        "Customer", "Original Entry ID", "Revision", "Previous Price", "New Price",
        "Changed By", "Changed At", "Reason", "Restored From Revision",
    ]
    sheet.append(headers)
    header_fill = PatternFill("solid", fgColor="113943")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left")
        cell.border = Border(bottom=Side(style="thin", color="D8C7AF"))
    for record in records:
        sheet.append([
            record.get("customer_name"), record.get("original_entry_id"), record.get("revision_number"),
            record.get("previous_price"), record.get("new_price"), record.get("changed_by"),
            record.get("changed_at"), record.get("change_reason"), record.get("restored_from_revision"),
        ])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column in (4, 5):
        for cell in sheet.iter_cols(min_col=column, max_col=column, min_row=2):
            for price_cell in cell:
                price_cell.number_format = "#,##0.00"
    widths = [30, 28, 12, 16, 16, 32, 24, 42, 24]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    return workbook

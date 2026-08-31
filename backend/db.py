import os
from pathlib import Path

from pymongo import MongoClient
from pymongo.uri_parser import parse_uri


DEFAULT_MONGODB_URI = "mongodb://localhost:27017/inventory_app"
ENV_FILE = Path(__file__).resolve().parent.parent / ".env"
_client = None
_database = None


def _load_local_env():
    """Load simple KEY=VALUE settings without overriding real environment variables."""
    try:
        lines = ENV_FILE.read_text(encoding="utf-8").splitlines()
    except OSError:
        return

    for raw_line in lines:
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {"\"", "'"}:
            value = value[1:-1]
        if key:
            os.environ.setdefault(key, value)


_load_local_env()


def get_mongodb_uri():
    return os.getenv("MONGODB_URI") or os.getenv("MONGO_URI") or DEFAULT_MONGODB_URI


def get_client():
    global _client

    if _client is None:
        _client = MongoClient(get_mongodb_uri(), serverSelectionTimeoutMS=5000)

    return _client


def get_database():
    global _database

    if _database is None:
        mongodb_uri = get_mongodb_uri()
        database_name = parse_uri(mongodb_uri).get("database") or "inventory_app"
        _database = get_client()[database_name]

    return _database


def get_inventory_collection():
    collection = get_database()["inventory_items"]
    legacy_key = [
        ("category", 1),
        ("brand", 1),
        ("type", 1),
        ("width", 1),
        ("height", 1),
        ("thickness", 1),
    ]
    for index_name, index_info in collection.index_information().items():
        if index_name == "_id_":
            continue
        if (
            index_name in {"inventory_item_identity_v2", "inventory_item_identity_v3"}
            or (index_info.get("key") == legacy_key and index_info.get("unique"))
        ):
            collection.drop_index(index_name)
    collection.create_index(
        [
            ("category", 1),
            ("brand", 1),
            ("type", 1),
            ("batch_roll_no", 1),
            ("width", 1),
            ("height", 1),
            ("thickness", 1),
            ("blanket_name", 1),
            ("nominal_width", 1),
            ("actual_width", 1),
            ("length_meters", 1),
            ("roll_no", 1),
            ("batch_no", 1),
            ("print_type", 1),
            ("storage_type", 1),
        ],
        unique=True,
        name="inventory_item_identity_v4",
    )
    return collection


def get_stock_logs_collection():
    collection = get_database()["stock_logs"]
    collection.create_index([("item_key", 1), ("changed_at", -1)])
    return collection


def get_users_collection():
    collection = get_database()["users"]
    collection.create_index([("email", 1)], unique=True, name="user_email_unique")
    collection.create_index([("role", 1)])
    return collection

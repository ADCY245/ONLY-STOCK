import os

from pymongo import MongoClient
from pymongo.uri_parser import parse_uri


DEFAULT_MONGODB_URI = "mongodb://localhost:27017/inventory_app"
_client = None
_database = None


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
        if index_info.get("key") == legacy_key and index_info.get("unique"):
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
        ],
        unique=True,
        name="inventory_item_identity_v2",
    )
    return collection


def get_stock_logs_collection():
    collection = get_database()["stock_logs"]
    collection.create_index([("item_key", 1), ("changed_at", -1)])
    return collection

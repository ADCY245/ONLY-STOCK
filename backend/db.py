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
    collection.create_index(
        [
            ("category", 1),
            ("brand", 1),
            ("type", 1),
            ("width", 1),
            ("height", 1),
            ("thickness", 1),
        ],
        unique=True,
    )
    return collection


def get_stock_logs_collection():
    collection = get_database()["stock_logs"]
    collection.create_index([("item_key", 1), ("changed_at", -1)])
    return collection

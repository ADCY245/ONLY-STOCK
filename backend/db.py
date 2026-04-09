import os

from pymongo import MongoClient
from pymongo.uri_parser import parse_uri


MONGODB_URI = os.getenv("MONGODB_URI", "mongodb://localhost:27017/inventory_app")
client = MongoClient(MONGODB_URI)


def get_database():
    database_name = parse_uri(MONGODB_URI).get("database") or "inventory_app"
    return client[database_name]


def get_inventory_collection():
    collection = get_database()["inventory_items"]
    collection.create_index(
        [("category", 1), ("brand", 1), ("type", 1), ("size", 1)],
        unique=True,
    )
    return collection


def get_stock_logs_collection():
    collection = get_database()["stock_logs"]
    collection.create_index([("item_key", 1), ("changed_at", -1)])
    return collection

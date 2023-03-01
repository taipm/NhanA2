from pymongo import MongoClient
from bson.objectid import ObjectId
import pandas as pd

class MongoDB:
    def __init__(self, db_name, xls_data_path, db_collection):
        self.client = MongoClient()
        self.db = self.client[db_name]
        self.collection = db_collection
        self.xls_data_path = xls_data_path 
    
    def create(self, data):
        if self.collection:
            data['sent_email'] = False
            result = self.collection.insert_one(data)
            return str(result.inserted_id)
        else:
            raise ValueError("Collection not set")
    
    def save_all(self):
        items = self.get_data()
        if len(items) > 0:
            count = 0
            for item in items:
                self.create(item)
                print(f'Upload to Db: {count}/{len(items)}')

    def read(self, query=None):
        if self.collection:
            if query:
                result = self.collection.find(query)
            else:
                result = self.collection.find()
            return [record for record in result]
        else:
            raise ValueError("Collection not set")
    
    def update(self, record_id, data):
        if self.collection:
            result = self.collection.update_one({'_id': ObjectId(record_id)}, {'$set': data})
            return result.modified_count
        else:
            raise ValueError("Collection not set")
    
    def delete(self, record_id):
        if self.collection:
            result = self.collection.delete_one({'_id': ObjectId(record_id)})
            return result.deleted_count
        else:
            raise ValueError("Collection not set")


from pymongo import MongoClient

class MongoDb:
    def __init__(self, host, port, username, password, database):
        self.client = MongoClient(host=host, port=port, username=username, password=password)
        self.db = self.client[database]

    def create(self, collection_name, document):
        collection = self.db[collection_name]
        result = collection.insert_one(document)
        return result.inserted_id

    def read(self, collection_name, filter=None):
        collection = self.db[collection_name]
        documents = collection.find(filter)
        return [doc for doc in documents]

    def update(self, collection_name, filter, update):
        collection = self.db[collection_name]
        result = collection.update_one(filter, {'$set': update})
        return result.modified_count

    def delete(self, collection_name, filter):
        collection = self.db[collection_name]
        result = collection.delete_one(filter)
        return result.deleted_count

import configparser

config = configparser.ConfigParser()
config.read('config.ini')

db_host = config['DATABASE']['host']
db_port = config['DATABASE']['port']
db_name = config['DATABASE']['name']

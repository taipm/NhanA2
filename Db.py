import pymongo
import pandas as pd



# data = getData(file_path='/Users/taipm/Documents/GitHub/NhanA2/Data/BINNIES/Images/CONG TY BINNIES UK LIMITED_NHAN VIEN_14.xlsm')
# print(data)
def save_to_mongodb(data):
    # Kết nối tới CSDL MongoDb
    client = pymongo.MongoClient("mongodb://localhost:27017/")
    db = client["mydatabase"]
    collection = db["mycollection"]
    
    # Thêm trường "Sent email" với giá trị False
    data['Sent email'] = False
    
    # Đổi tên trường từ tiếng Việt sang tiếng Anh
    data = {
        'Contract holder': data['Chủ hợp đồng:'],
        'Contract number': data['Số hợp đồng:'],
        'Validity': data['Hiệu lực:'],
        'Beneficiary name': data['Họ tên NĐBH:'],
        'Year of birth': data['Năm sinh:'],
        'ID card/CCCD': data['CMND/CCCD:'],
        'Email': data['EMAIL:'],
        'Sent email': data['Sent email']
    }
    
    # Lưu dữ liệu vào CSDL
    result = collection.insert_one(data)
    
    # In thông tin về dữ liệu đã được lưu
    print("Data saved with ID:", result.inserted_id)

from pymongo import MongoClient
from bson.objectid import ObjectId

class MongoDB:
    def __init__(self, db_name, xls_data_path, db_collection):
        self.client = MongoClient()
        self.db = self.client[db_name]
        self.collection = db_collection
        self.xls_data_path = xls_data_path 
    
    def get_data(self, sheetName = 'Data'):
        fileName = self.xls_data_path
        data = pd.read_excel(fileName, sheet_name=sheetName, index_col=None, header=None)
        data.dropna(inplace=True)
        dict_list = []

        for row in data.itertuples():
            dict = {}
            dict[row[1]] = row[2]
            dict_list.append(dict)

        return dict_list
    
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

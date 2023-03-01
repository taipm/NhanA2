import pandas as pd
from pymongo import MongoClient
import certifi
import ssl
ssl._create_default_https_context = ssl._create_unverified_context


def get_data(xls_data_path, sheetName = 'Data'):
  data = pd.read_excel(xls_data_path, sheet_name=sheetName, index_col=None, header=None)
  data.dropna(inplace=True)
  dict_list = []
  n_rows = 7
  for i in range(0, len(data), n_rows):
    chunk = data[i:i+n_rows]
    dict = {}
    for row in chunk.itertuples():
      dict[row[1]] = row[2]
      
    dict_list.append(dict)
  return dict_list

def save_to_mongodb(data):
    client = MongoClient("mongodb+srv://taipm:OAMOHMEC8CPUHoz2@cluster0.nskndlz.mongodb.net/?retryWrites=true&w=majority",tlsCAFile=certifi.where())
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
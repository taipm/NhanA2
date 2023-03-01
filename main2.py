import configparser
from datetime import datetime
from mongodb import MongoDb

config = configparser.ConfigParser()
config.read('config.ini')

host = config['database']['host']
port = int(config['database']['port'])
username = config['database']['username']
password = config['database']['password']
database = config['database']['database']

db = MongoDb(host, port, username, password, database)

# Thêm một document mới vào collection
document = {
    'name': 'John Doe',
    'email': 'johndoe@example.com',
    'created_at': datetime.now()
}
db.create('users', document)

# Lấy ra tất cả các document trong collection
documents = db.read('users')
for doc in documents:
    print(doc)

# Cập nhật một document trong collection
filter = {'name':

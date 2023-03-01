from ctypes.wintypes import RGB
from time import sleep
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import datetime
from datetime import datetime
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from Email import *

def getData(file_path):
  fileName = file_path
  data = pd.read_excel(fileName, sheet_name='Data', index_col=None, header=None)
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

def add_table_to_image(img_path, new_img_path, data, fontsize=20):
    image = Image.open(img_path)
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype("Arial.ttf", 35)

    x, y = 100, 100 # Tọa độ góc trái trên của bảng dữ liệu
    #TAIPM: manual adjust
    x = x + 30 + 100
    y = y + 50 + 150
    padding = 25 # Khoảng cách giữa các ô
    khoang_cach = 45
    for i, (key, value) in enumerate(data.items()):
        color = 'black'
        draw.text((x, y + padding + i*khoang_cach), str(key), font=font, fill= color)
        draw.text((x + 280, y + padding + i*khoang_cach), str(value), font=font, fill= color)

    image.save(new_img_path)


def get_img(item):
  file_name = f"{str(item['CMND/CCCD:'])}_{str(item['EMAIL:']).replace('@','_').replace('.','_')}_{item['Họ tên NĐBH:']}_{str(item['Năm sinh:']).replace('/','_').replace('-','_').replace(' 00:00:00','')}"
  return file_name

def make_images(file_path, file_mau, img_folder):
    #HÌNH MẪU: 
    img_path = file_mau
    fontsize = 28
    items = getData(file_path=file_path)
    #THAY Ở ĐÂY
    print(f'Số lượng bản ghi: {len(items)}')
    count_img = 0
    for item in items:#[0:1]: #sửa ở đây
        data = item        
        img_name = get_img(item=data)
        new_img_path = f"{img_folder}/{img_name}.jpg"
        datetime_str = str(item['Năm sinh:']).replace('/','-')
        try:
          datetime_obj = datetime.strptime(datetime_str, "%Y-%m-%d %H:%M:%S")
        except :
          datetime_obj = datetime.strptime(datetime_str, "%d-%m-%Y")

        date_str = datetime_obj.strftime("%d-%m-%Y")
        
        data = {
            'Chủ hợp đồng': item['Chủ hợp đồng:'],
            'Số hợp đồng': item['Số hợp đồng:'],
            'Hiệu lực': item['Hiệu lực:'],
            'Họ tên NĐBH': item['Họ tên NĐBH:'],
            'Năm sinh': date_str,
            'CMND/CCCD': item['CMND/CCCD:']
        }
        
        add_table_to_image(img_path, new_img_path, data, fontsize)
        count_img += 1
    return count_img


Erros = [] 
def sendEmail(item_data):
    username = "ntnhan@baominh.com.vn"
    password = "thuantienPHAT30"
    mail_from = "ntnhan@baominh.com.vn"
  
    if len(item_data['Email:']) == 0 and '@' in item_data:
        print(f'Lỗi dữ liệu: {item_data}')
        return

    #mail_to = 'taipm.vn@gmail.com'
    mail_to = item_data['Email:']
    print(item_data['Email:'])
    mail_title = getTitle(custName = item_data['Họ tên NĐBH:'])
    mail_body = getBody()

    mimemsg = MIMEMultipart()
    mimemsg['From'] = mail_from
    mimemsg['To'] = mail_to
    mimemsg['Subject'] = mail_title
    mimemsg.attach(MIMEText(mail_body, 'html'))    
    img_path = f'{img_folder}/{get_img(item_data)}.jpg'
    print(img_path)
    with open(img_path, 'rb') as f:
        img_data = f.read()
    img = MIMEImage(img_data)
    img.add_header('Content-Disposition', 'attachment', filename=img_path)
    mimemsg.attach(img)

    connection = smtplib.SMTP(host='smtp.office365.com', port=587)
    connection.starttls()
    connection.login(username, password)
    try:
        connection.send_message(mimemsg)
    except:
        Erros.append(item_data['Email:'])
        print(f'Lỗi: {item_data}')
    connection.quit()

files = [
   #'/Users/taipm/Documents/GitHub/live-capture/NhanA2/SAVILLS/Data/DANH SACH BAO HIEM SAVILLS HCM_HO.xlsm',
   #'/Users/taipm/Documents/GitHub/live-capture/NhanA2/SAVILLS/Data/PL3_DANH SACH BAO HIEM SAVILLS HCM_PM 1-500.xlsm',
   #'/Users/taipm/Documents/GitHub/live-capture/NhanA2/SAVILLS/Data/PL3_DANH SACH BAO HIEM SAVILLS HCM_PM 501-1000.xlsm',
   #'/Users/taipm/Documents/GitHub/live-capture/NhanA2/SAVILLS/Data/PL3_DANH SACH BAO HIEM SAVILLS HCM_PM 1001-1500.xlsm',
   #'/Users/taipm/Documents/GitHub/live-capture/NhanA2/SAVILLS/Data/PL3_DANH SACH BAO HIEM SAVILLS HCM_PM 1501 -1913.xlsm',
   '/Users/taipm/Documents/GitHub/live-capture/NhanA2/BINNIES/CONG TY BINNIES UK LIMITED_NHAN VIEN_15.xlsm'   
]

#file_path = '/Users/taipm/Documents/GitHub/live-capture/NhanA2/NamLong/NAM LONG_ NGUOI THAN.xlsm'
#data_folder = '/Users/taipm/Documents/GitHub/live-capture/NhanA2/SAVILLS/Data'
file_mau =  '/Users/taipm/Documents/GitHub/live-capture/NhanA2/BINNIES/MAU THE_KO QUA MG_48.jpg'
img_folder = '/Users/taipm/Documents/GitHub/live-capture/NhanA2/BINNIES/Images'


#STEP 0: XÓA DỮ LIỆU
def CleanImages(img_folder):
  import os, shutil
  folder = img_folder
  for filename in os.listdir(folder):
      file_path = os.path.join(folder, filename)
      try:
          if os.path.isfile(file_path) or os.path.islink(file_path):
              os.unlink(file_path)
          elif os.path.isdir(file_path):
              shutil.rmtree(file_path)
      except Exception as e:
          print('Failed to delete %s. Reason: %s' % (file_path, e))


#STEP 01: Sinh ảnh cho toàn bộ dữ liệu
def MakeImages():
  count_img = 0
  for file in files:
      items = getData(file_path=file)
      print(len(items))
      count_img += make_images(file_path=file, file_mau=file_mau, img_folder=img_folder)
  print(f'Đã sinh tổng cộng: {count_img} ảnh tại thư mục {img_folder}')

#STEP 2: Gửi email
def NotifyEmail():
  count_email = 0
  for file in files:
      items = getData(file_path=file)
      print(f'Đang gửi email cho file : {file}')
      print(f'Tổng số bản ghi: {len(items)}')

      count = 0
      for item in items:#[0:1]:        
          print(item)        
          print(f"Đang gửi: {count}/{len(items)}")
          sendEmail(item)
          count_email = count_email + 1
          count = count + 1
          if count_email % 10 == 0:
            sleep(3)
            
  print(f'Đã gửi tổng cộng: {count_email}')

  print(f'LỖI: {len(Erros)}')    
  for e in Erros:
    print(f'{e}')

CleanImages(img_folder=img_folder)
MakeImages()

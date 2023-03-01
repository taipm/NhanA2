from ctypes.wintypes import RGB
from time import sleep
from PIL import Image, ImageDraw, ImageFont
import datetime
from datetime import datetime
from DataProcessor import get_data
from Email import *
import glob
import os
#import configparser

def generate_Img(img_path, new_img_path, data, fontsize, font_color):
    image = Image.open(img_path)
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(font= "Arial.ttf", size=fontsize)

    # x, y = 100, 100 # Tọa độ góc trái trên của bảng dữ liệu
    # #TAIPM: manual adjust
    # x = x + 30 + 100
    # y = y + 50 + 150
    # padding = 25 # Khoảng cách giữa các ô
    #khoang_cach = 45
    for i, (key, value) in enumerate(data.items()):
        color = font_color
        draw.text((x, y + padding + i*line_space), str(key), font=font, fill= color)
        draw.text((x + 280, y + padding + i*line_space), str(value), font=font, fill= color)

    image.save(new_img_path)


def get_img(item):
  
  file_name = f"{str(item['CMND/CCCD:'])}"+\
        f"_{str(item['EMAIL:']).replace('@','_').replace('.','_')}"+\
        f"_{item['Họ tên NĐBH:']}"+\
        f"_{str(item['Năm sinh:']).replace('/','_').replace('-','_').replace(' 00:00:00','')}"
  
  return file_name

def make_images(file_path, file_mau, img_folder, font_size=30):
    img_path = file_mau
    items = get_data(xls_data_path=file_path)
    print(f'Số lượng bản ghi: {len(items)}')
    count_img = 0
    for item in items:#[0:1]: #sửa ở đây
        #data = item
        img_name = get_img(item=item)
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
        print(data)
        generate_Img(img_path, new_img_path, data, fontsize=font_size, font_color=font_color)
        count_img += 1
    return count_img




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
def MakeImages(xlsm_files, card_path):
  count_img = 0
  for file in xlsm_files:
      print(file)
      items = get_data(xls_data_path=file)
      print(len(items))
      for item in items:
         print(item)
      count_img += make_images(file_path=file, file_mau=card_path, img_folder=img_folder)
  print(f'Đã sinh tổng cộng: {count_img} ảnh tại thư mục {img_folder}')

#STEP 2: Gửi email
def NotifyEmail(xlsm_files):
    count_email = 0
    for file in xlsm_files:
        items = get_data(xls_data_path=file)
        print(f'Đang gửi email cho file : {file}')
        print(f'Tổng số bản ghi: {len(items)}')

        count = 0
        for item in items:#[0:1]:        
            print(item)        
            print(f"Đang gửi: {count}/{len(items)}")
            email_title = getTitle(custName=item['Họ tên NĐBH:'])
            email_to = item['EMAIL:']
            img_path = f'{img_folder}/{get_img(item=item)}.jpg'

            result = sendEmail(title=email_title,to_email=email_to,img_path=img_path)
            if len(str(result)) > 0:
                Errors.append(result)

            count_email = count_email + 1
            count = count + 1
            if count_email % 10 == 0:
                sleep(3)
                
    print(f'Đã gửi tổng cộng: {count_email}')
    print(f'LỖI: {len(Erros)}')    
    for e in Erros:
        print(f'{e}')

def create_folders():
    if not os.path.exists(img_folder):
        os.makedirs(img_folder)
    if not os.path.exists(card_path):
        with open(card_path, 'w'): 
           print(f'Chưa có file mẫu: .jpg')


companyName = 'BINNIES'
card_name = 'MAU THE_KO QUA MG_48.jpg'
font_size = 30
font_color = 'Black'
line_space = 40
x, y = 100, 100 # Tọa độ góc trái trên của bảng dữ liệu
x = x + 30 + 100
y = y + 50 + 150
padding = 25 # Khoảng cách giữa các ô

current_dir = os.getcwd()
img_folder = f'{current_dir}/Data/{companyName}/images'
card_path = f'{current_dir}/Data/{companyName}/{card_name}'
data_folder = f'{current_dir}/Data/{companyName}'
Errors = [] 
xlsm_files = glob.glob(f'{data_folder}/*.xlsm')

create_folders()
#CleanImages(img_folder=img_folder)
MakeImages(xlsm_files=xlsm_files, card_path=card_path)
NotifyEmail(xlsm_files=xlsm_files)
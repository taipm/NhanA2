import configparser
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def getTitle(custName):
  title = 'THẺ BẢO LÃNH ĐIỆN TỬ - CÔNG TY BẢO MINH SÀI GÒN - ' + custName.upper()
  return title
  
def getBody():
  html = '''<html>
              <head></head>
              <h3>Kính gửi: QUÝ KHÁCH HÀNG</h3>
              <body>
                <p>Công ty Bảo Minh Sài Gòn trân trọng gửi đến Quý khách hàng thẻ bảo lãnh điện tử, vui lòng xem file đính kèm.</p>
                <h4><i>Dear Valued Customer,</i></h4>
                <p><i>Bao Minh Sai Gon would like to send you an electronic insurance card, please see the attached file.</i></p>
                <p>Thanks &amp; Best Regards,</p>
                <p>For And On Behalf Of Mr Tran Nam Thanh, G/D</p>
                <p>==================================================<p>
                <p>Nguyen Thi Nhan (Mrs)</p>
                <p>Deputy Manager</p>
                <p>Flr 5, Bao Minh Building, 217 Nam Ky Khoi Nghia, Ward Vo Thi Sau, Dist 3, HCMC</p>
                <p>BAO MINH SAI GON</p>
                <p>Tel: (+84.28) 39326262</p>
                <p>Website: www.baominh.com.vn</p>
              </body>
            </html>'''
  return html

config = configparser.ConfigParser()
config.read('config.ini')

email_user = config['Email']['username']
email_pwd = config['Email']['password']
email_from = config['Email']['mail_from']

def sendEmail(title, to_email, img_path):
    username = email_user
    password = email_pwd
    mail_from = email_from
    mail_to = 'taipm.vn@gmail.com'
    #mail_to = item_data['EMAIL:']
    mail_to = to_email
    print(mail_to)

    if len(mail_to) == 0 and '@' in mail_to:
        print(f'Lỗi dữ liệu: {mail_to}')
        return

    mail_title = title
    mail_body = getBody()

    mimemsg = MIMEMultipart()
    mimemsg['From'] = mail_from
    mimemsg['To'] = mail_to
    mimemsg['Subject'] = mail_title
    mimemsg.attach(MIMEText(mail_body, 'html'))
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
        connection.quit()
        return
    except:
        print(f'Lỗi: {mail_to}')
        connection.quit()
        return mail_to
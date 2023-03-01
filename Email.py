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
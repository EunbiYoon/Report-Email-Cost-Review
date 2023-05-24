import pandas as pd
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication 
from table import this_week, F_BPAE_html, T_BPAE_html, D_BPAE_html, F_PACE_html, T_PACE_html, D_PACE_html, FBI_html, TBI_html, DBI_html, FPI_html, TPI_html, DPI_html
from graph import save_path,today_date
import os

#html - table
server = smtplib.SMTP('lgekrhqmh01.lge.com:25')
server.ehlo()


#메일 내용 구성
msg=MIMEMultipart()

# 수신자 발신자 지정
msg['From']='eunbi1.yoon@lge.com'
msg['To']='catalina.manolache@lge.com'
# msg['To']='iggeun.kwon@lge.com'
# msg['Cc']='janine.williams@lge.com, karina1.beveridge@lge.com, kitae3.park@lge.com, soyoung1.an@lge.com, soyoon1.kim@lge.com, wolyong.ha@lge.com, grace.hwang@lge.com, tg.kim@lge.com, seongju.yu@lge.com, minhyoung.sun@lge.com, jongseop.kim@lge.com, richard.song@lge.com, gilnam.lee@lge.com, jacey.jung@lge.com, 312718@lge.com'
msg['Bcc']='eunbi1.yoon@lge.com'

#Subject 꾸미기
msg['Subject']='Cost Review Report '+this_week
# msg.attach(MIMEText('<h4 style="font-family:sans-serif; font-weight:500">Dear all,<br/>I would like to share this week cost review report and detailed informations are in below website.<br/>You can access website from CloudPC or LG wifi for security purpose. <a href="http://10.225.2.85">http://10.225.2.85</a></h4>','html'))
msg.attach(MIMEText('<h3 style="font-family:sans-serif;">Dear Catalina,</h3><h4 style="font-family:sans-serif; font-weight:500">sharing you to updated version, <br/>I just want to send this email for this week if the data is correct, I can update more on next week.</h4>','html'))


# graph file read
with open(save_path+"FL_BPA_Entity"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image1 = MIMEImage(img_data, name=os.path.basename("FL_BPA_Entity"+today_date+'.png'))

with open(save_path+"TL_BPA_Entity"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image2 = MIMEImage(img_data, name=os.path.basename("TL_BPA_Entity"+today_date+'.png'))

with open(save_path+"DR_BPA_Entity"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image3 = MIMEImage(img_data, name=os.path.basename("DR_BPA_Entity"+today_date+'.png'))


with open(save_path+"FL_PAC_Entity"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image4 = MIMEImage(img_data, name=os.path.basename("FL_PAC_Entity"+today_date+'.png'))

with open(save_path+"TL_PAC_Entity"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image5 = MIMEImage(img_data, name=os.path.basename("TL_PAC_Entity"+today_date+'.png'))

with open(save_path+"DR_PAC_Entity"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image6 = MIMEImage(img_data, name=os.path.basename("DR_PAC_Entity"+today_date+'.png'))


# html table attach
F_BPAE_attach = MIMEText(F_BPAE_html, "html")
T_BPAE_attach = MIMEText(T_BPAE_html, "html")
D_BPAE_attach = MIMEText(D_BPAE_html, "html")

F_PACE_attach = MIMEText(F_PACE_html, "html")
T_PACE_attach = MIMEText(T_PACE_html, "html")
D_PACE_attach = MIMEText(D_PACE_html, "html")

FBI_attach = MIMEText(FBI_html, "html")
TBI_attach = MIMEText(TBI_html, "html")
DBI_attach = MIMEText(DBI_html, "html")

FPI_attach = MIMEText(FPI_html, "html")
TPI_attach = MIMEText(TPI_html, "html")
DPI_attach = MIMEText(DPI_html, "html")

msg.attach(MIMEText('<br/><h3 style="font-family:sans-serif;">Front Loader BPA Entity Trend</h3>','html'))
msg.attach(image1)
msg.attach(F_BPAE_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(FBI_attach)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:sans-serif;">Top Loader BPA Entity Trend</h3>','html'))
msg.attach(image2)
msg.attach(T_BPAE_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(TBI_attach)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:sans-serif;">Dryer BPA Entity Trend</h3>','html'))
msg.attach(image3)
msg.attach(D_BPAE_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(DBI_attach)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:sans-serif;">Front Loader PAC Entity Trend</h3>','html'))
msg.attach(image4)
msg.attach(F_PACE_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(FPI_attach)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:sans-serif;">Top Loader PAC Entity Trend</h3>','html'))
msg.attach(image5)
msg.attach(T_PACE_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(TPI_attach)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:sans-serif;">Dryer PAC Entity Trend</h3>','html'))
msg.attach(image6)
msg.attach(D_PACE_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(DPI_attach)


#첨부 파일1
etcFileName='Cost Review_0526.xlsx'
with open("C:/Users/RnD Workstation/Documents/CostReview/0526/Cost Review_0526.xlsx", 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)


#메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")
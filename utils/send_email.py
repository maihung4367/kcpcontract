
from datetime import datetime
from django.template.loader import get_template
from django.core.mail import EmailMessage
from django.conf import settings
from user.models import User,Profile
import requests
from excelExtract.models import pdfFile
def send_noti_to_partner_sign_by_email(ct,account,list_system_link_file_pdf, customer_email):
	
	subject = "KCV_THÔNG BÁO CHƯƠNG TRÌNH {} THÁNG {} {} ".format(str(ct).upper(),datetime.now().strftime("%m.%Y"),str(account).upper())
	html_message = get_template("template_email.html").render({"ct":ct})

	msg = EmailMessage(subject,html_message,settings.EMAIL_HOST_USER,to=customer_email)
	msg.content_subtype = "html"

	for linkfile in list_system_link_file_pdf:
		try:
			msg.attach('{}'.format(linkfile.replace(settings.URL+"/"+"documents/slavefiles/","")), requests.get(linkfile,allow_redirects=True).content)
		except:
			pass
	msg.send()
def send_noti_to_partner_sign_by_email2(list_system_link_file_pdf, customer_email):
	userList=User.objects.filter(is_signer=True)
	user=userList[0]
	userFullname=Profile.objects.get(user=user).full_name
	numberunsignepdfs=len(pdfFile.objects.filter(signed=False,sended=False,confirmed=True))
	subject = " KCV_THÔNG BÁO CHƯƠNG TRÌNH CẦN KÝ THÁNG {}".format(datetime.now().strftime("%m.%Y"))
	html_message = get_template("template_email2.html").render({"customer_email":customer_email,"numberunsignepdfs":numberunsignepdfs,"fullname":userFullname})

	msg = EmailMessage(subject,html_message,settings.EMAIL_HOST_USER,to=[customer_email,])
	msg.content_subtype = "html"

	
	msg.send()
def send_noti_to_confirmer(list_system_link_file_pdf, customer_emails):

	subject = " KCV_THÔNG BÁO CHƯƠNG TRÌNH CẦN XÁC NHẬN THÁNG {}".format(datetime.now().strftime("%m.%Y"))
	html_message = get_template("notifyConfirm.html").render({"customer_email":customer_emails})

	msg = EmailMessage(subject,html_message,settings.EMAIL_HOST_USER,to=customer_emails)
	msg.content_subtype = "html"

	
	msg.send()

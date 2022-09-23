
from django.template.loader import get_template
from django.core.mail import EmailMessage
from django.conf import settings
import requests
from excelExtract.models import pdfFile
def send_noti_to_partner_sign_by_email(list_system_link_file_pdf, customer_email):
	
	subject = " {} ".format(list_system_link_file_pdf[0].replace(settings.URL+"/"+"/documents/slavefiles/",""))
	html_message = get_template("template_email.html").render({"customer_email":customer_email,"file_pdf":list_system_link_file_pdf[0]})

	msg = EmailMessage(subject,html_message,settings.EMAIL_HOST_USER,to=customer_email)
	msg.content_subtype = "html"

	for linkfile in list_system_link_file_pdf:
		try:
			msg.attach('{}'.format(linkfile.replace(settings.URL+"/"+"documents/slavefiles/","")), requests.get(linkfile,allow_redirects=True).content)
		except:
			pass
	msg.send()
def send_noti_to_partner_sign_by_email2(list_system_link_file_pdf, customer_email):
	numberunsignepdfs=len(pdfFile.objects.filter(signed=False,sended=False,confirmed=True))
	subject = " Có {} văn bản cần kí gửi ".format(numberunsignepdfs)
	html_message = get_template("template_email2.html").render({"customer_email":customer_email,"file_pdf":list_system_link_file_pdf[0]})

	msg = EmailMessage(subject,html_message,settings.EMAIL_HOST_USER,to=customer_email)
	msg.content_subtype = "html"

	for linkfile in list_system_link_file_pdf:
		try:
			msg.attach('{}'.format(linkfile.replace(settings.URL+"/"+"documents/slavefiles/","")), requests.get(linkfile,allow_redirects=True).content)
		except:
			pass
	msg.send()
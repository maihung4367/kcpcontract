
from django.template.loader import get_template
from django.core.mail import EmailMessage
from django.conf import settings
import requests

def send_noti_to_partner_sign_by_email(file_pdf, customer_email):
	subject = "PVS - You are assigned to {} on P-sign".format(file_pdf)
	html_message = get_template("template_email.html").render({"customer_email":customer_email,"file_pdf":file_pdf})

	msg = EmailMessage(subject,html_message,settings.EMAIL_HOST_USER,to=[customer_email,])
	msg.content_subtype = "html"
	msg.attach('HD_pdf.pdf', requests.get(file_pdf,allow_redirects=True).content)

	msg.send()
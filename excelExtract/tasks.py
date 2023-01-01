from celery import  Celery
from celery import shared_task
from .telegram import kc_project_alert
from .auto_email_report import send_report
@shared_task
def email_report_and_telegram_annouce():
	send_report()
	kc_project_alert()

@shared_task
def add(x,y):
	return x+y
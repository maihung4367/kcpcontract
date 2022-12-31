from celery import  Celery
from celery import shared_task
from datetime import datetime, timedelta
from .models import pdfFile
from django.http import HttpResponse
import openpyxl
from openpyxl.styles import Alignment,NamedStyle
from openpyxl.writer.excel import save_virtual_workbook
def is_leap_year(year): 
	if year % 100 == 0:
		return year % 100 == 0

	return year % 4 == 0

def get_lapse():
	last_month = datetime.today().month - 1

	print(last_month)

	current_year = datetime.today().year

	#is last month a month with 30 days?
	if last_month in [9, 4, 6, 11]:
		lapse = 30

	#is last month a month with 31 days?
	elif last_month in [1, 3, 5, 7, 8, 10, 12]:
		lapse = 31

	#is last month February?
	else:
		if is_leap_year(current_year):
			lapse = 29
		else:
			lapse = 30

	return lapse


# @shared_task(name="test")
# def test_celery():
# 	last_month_filter = datetime.today() - timedelta(days=get_lapse())
# 	print(get_lapse())
# 	print(last_month_filter.replace(minute=0, hour=0, second=0, microsecond=0))
# 	print(datetime.today().replace(minute=0, hour=0, second=0, microsecond=0))
# @shared_task(name="auto_clean_trash")
# def auto_clean_trash():
# 	trashs = pdfFile.objects.filter(is_deleted=True)
# 	trashs.delete()
# 	print(trashs)
# 	print(len(trashs))


@shared_task
def add(x,y):
	return x+y

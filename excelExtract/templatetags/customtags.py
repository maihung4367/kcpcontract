from django import template
from excelExtract.models import document,excel
register =template.Library() 

@register.filter(name="nameFileFilter")
def nameFileFilter(value,*args):
	value=str(value)
	return value.replace("documents/","")
@register.filter(name="subNameFileFilter")
def subNameFileFilter(value,*args):
	value=str(value)
	return value.replace("documents/slavefiles/","")

@register.filter(name="accountFilter")
def accountFilter(value):
	# listloaict=[]
	listaccount=[]
	file=document.objects.get(pk=int(value))
	for f in excel.objects.filter(filename=file):
		if f.account not in listaccount:
			listaccount.append(f.account)
		# if f.loaiCt not in listloaict:
		# 	listloaict.append(f.loaiCt)
	return 	listaccount

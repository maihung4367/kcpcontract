from django import template
from excelExtract.models import document,excel,pdfFile
from user.models import Profile
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
@register.filter(name="confirmedValue")
def confirmedValue(value):
	confirmedValue=""
	if value==True:
		confirmedValue="Đã xác nhận"
	return 	confirmedValue
@register.filter(name="signedValue")
def signedValue(value):
	if value==True:
		signedValue="Đã gửi"
	else:
		signedValue="Chưa gửi"
	return 	signedValue
@register.filter(name="checkSlaveFile")
def scheckSlaveFile(excelfile):
	if pdfFile.objects.filter(masterFile=excelfile,confirmed=True).exists():
		
		return False
	else:
		return True


@register.simple_tag()
def get_info_profile(user):
	profile = Profile.objects.get(user=user)
	return {"profile":profile}
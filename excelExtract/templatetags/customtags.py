from django import template
from excelExtract.models import document,excel, excelAccount,pdfFile
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
@register.filter(name="staffProfileEmail")
def staffProfileEmail(staff):
	email=Profile.objects.get(user=staff).email
	return email

@register.filter(name="staffAccount")
def staffAccount(staff):
	staffProfile=Profile.objects.get(user=staff)
	print(staffProfile)
	if excelAccount.objects.filter(responsibleBy=staffProfile).exists():
		account=excelAccount.objects.filter(responsibleBy=staffProfile)
		return account
	else:
		return None

@register.filter(name="staffUser")
def staffUser(staff):
	staffProfile=Profile.objects.get(user=staff)
	print(staffProfile)
	if excelAccount.objects.filter(responsibleBy=staffProfile).exists():
		account=excelAccount.objects.filter(responsibleBy=staffProfile)
		return account
	else:
		return None
@register.simple_tag()
def get_info_profile(user):
	profile = Profile.objects.get(user=user)
	return {"profile":profile}
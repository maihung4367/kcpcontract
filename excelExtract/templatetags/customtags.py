from django import template
from excelExtract.models import document,excel, excelAccount,pdfFile
from user.models import Profile
import re
from datetime import datetime
register =template.Library() 

@register.filter(name="nameFileFilter")
def nameFileFilter(value,*args):
	value=str(value)
	return value.replace("documents/","")
@register.filter(name="subNameFileFilter")
def subNameFileFilter(value,*args):
	value=str(value)
	return value.replace("documents/slavefiles/","")
@register.filter(name="short_name")
def short_name(value,*args):
	value=str(value).replace("documents/slavefiles/","")
	match = re.search(r'\d{4}-\d{2}-\d{2}', value)	
	# print(value,match)
	try:
		date = datetime.strptime(match.group(), '%Y-%m-%d').date()
		index=value.find(str(date))
		lenIndex=len(str(date))
		name=value[0:index+lenIndex]+ "...pdf"
		
		if len(value[0:index+lenIndex]) >= 27:
			name=value[0:26] + "...pdf"
		# print(name,len(name))
		return name
	except:
		return value
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
		
		return True
	else:
		return False
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
# ---------------------------------------------------------------------------------
@register.filter(name="email_filter_innerhtml")
def email_filter_innerhtml(email):
	string=str(email)
	for i,char in enumerate(string):
		if char == "@":
			name_of_string=string[0:i]
			return 	name_of_string

@register.filter(name="split_loaict_string")
def split_loaict_string(string):
	loaict_list=[]
	for f in string.split(','):
		loaict_list.append(f)
	return loaict_list
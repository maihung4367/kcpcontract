from urllib import response
from django.shortcuts import render,redirect
import sys
from rest_framework.decorators import api_view
from excelExtract.forms import uploadDocumentForm
from excelExtract.models import document, excelAccount,pdfFile, excel
from . import excelExtract
from user.models import Profile
from django.db import transaction
from rest_framework import status
from django.conf import settings
from django.http import HttpResponse
from rest_framework.response import Response
from django.views.decorators.clickjacking import xframe_options_sameorigin
from utils import send_email
from datetime import datetime
import requests
import json
from django.core.files import File
from django.db.models import Q
# Create your views here.
@xframe_options_sameorigin
def kcToolPage(request):
	if request.user.is_authenticated:
		form=uploadDocumentForm()
		files=document.objects.all()
		pdffiles=pdfFile.objects.all().order_by("-id")
		demoPdfFiles=pdfFile.objects.last()
		numberUnsignepdfs=len(pdfFile.objects.filter(confirmed=False))
		if request.method=='POST':
			with transaction.atomic():
				try:
					file=request.FILES['document']
					print(file)
					if file:
						try:
							user = request.user
							profile = Profile.objects.get(user=user)
							excelExtract.importDataExcel(file, user=profile)
						except:
							excelExtract.importDataExcel(file)
				except:
					pass
		return render(request,"KCtool/KCtool.html",{"form":form,"files":files,"pdffiles":pdffiles,"demoPdfFiles":demoPdfFiles,"numberUnsignepdfs":numberUnsignepdfs, "active_id":1})
	else :
		return HttpResponse("not authen")


def waitConfirmDoc(request):
	if request.user.is_authenticated:
		user=request.user
		numberunconfirmpdfs=len(pdfFile.objects.filter(confirmed=False))
		unconfirmpdfs=pdfFile.objects.filter(confirmed=False).order_by("-id")
		accountList=excelAccount.objects.all()
		profile=Profile.objects.get(user=user)
		listaccount=excelAccount.objects.filter(responsibleBy=profile)
		return render(request,"KCtool/waitingsigndoc.html",{"numberunconfirmpdfs":numberunconfirmpdfs,"unconfirmpdfs":unconfirmpdfs, "URL":settings.URL, "active_id":2,"accountList":accountList,"user":user,"listaccount":listaccount})
	else :
		return HttpResponse("not authen")
def signedDoc(request):
	if request.user.is_authenticated:
		user=request.user
		numbersignepdfs=len(pdfFile.objects.filter(signed=True))
		
		accountList=excelAccount.objects.all()
		if request.GET.get("key_word",""):
			key_word = request.GET.get("key_word")
			pdfs=pdfFile.objects.filter(confirmed=True,slaveFile__icontains=key_word).order_by("sended")
		else:
			pdfs=pdfFile.objects.filter(confirmed=True).order_by("sended")
		
		if request.GET.get("account",None):
			account=excelAccount.objects.filter(account=request.GET.get("account"))[0]
			print(account.pk)
			if request.GET.get("key_word",""):
				key_word = request.GET.get("key_word")
				pdfs = pdfFile.objects.filter(confirmed=True,account=account,slaveFile__icontains=key_word).order_by("sended")
			else:
				pdfs = pdfFile.objects.filter(confirmed=True,account=account).order_by("sended")
			
			
			
		return render(request,"KCtool/signedDoc.html",{"numbersignepdfs":numbersignepdfs,"pdfs":pdfs, "active_id":3,"user":user,"accountList":accountList,"key_word":request.GET.get("key_word",""),"account":request.GET.get("account",None)})

	else :
		return HttpResponse("not authen")
# def excelToListPdfs(request):  
#             return Response(request,"KCtool/KCTool.html")
@api_view(["POST"])
def getIdList(request):
	if request.method=='POST':
		print(request.data)
		# loaict=request.data['loaict']
		# listId=request.POST.getlist('file')
		# for  f in listId:
		#     loaiAccount=request.data['fileID{}'.format(f)]
		#     excelExtract.exportFiles(loaict=loaict,fileID=f,loaiAccount=loaiAccount)  
		return redirect("KCTool:kcToolPage")
	

@api_view(["POST"])
def create_pdf(request):
	try:
		with transaction.atomic():
			print(request.data)
			id_excel = int(request.data["pk_excel"])
			values_category = request.data["values_category"]
			values_account = request.data["values_account"]
			print(values_account)
			print(values_category)
			user=request.user
			profile = Profile.objects.get(user=user)
			annouce=excelExtract.exportFiles(loaict=values_category,fileID=id_excel,loaiAccount=values_account,user=profile) 
			print(annouce)
			return Response({"annouce":annouce}, status=status.HTTP_200_OK)
	except:
		return Response(status=status.HTTP_500_INTERNAL_SERVER_ERROR)


@api_view(["POST"])
def getListAccount(request):
	# try:
		arr_listaccounts=[]
		arr_loaiCt = []
		list_accounts = ""
		list_loaiCt = ""
		pk = request.data["pk_excel"]
		file=document.objects.get(pk=int(pk))
		list_accounts +=f'''
					<option group="list_accounts_group" value="all" data-badge="">All</option>
				'''
		list_loaiCt+= f'''
					<option group="list_LoaiCT_group" value="all" data-badge="">All</option>
				'''
		for f in excel.objects.filter(filename=file):
			if f.account not in arr_listaccounts:

				arr_listaccounts.append(f.account)
				list_accounts += f'''
					<option group="list_accounts_group" value="{f.account}" data-badge="">{f.account}</option>
				'''
			if f.loaiCt not in arr_loaiCt:
				arr_loaiCt.append(f.loaiCt)
				list_loaiCt += f'''
					<option group="list_LoaiCT_group" value="{f.loaiCt}" data-badge="">{f.loaiCt}</option>
				'''
		
		return Response({"Status":"Success", "list_accounts": list_accounts, "list_loaiCt":list_loaiCt},  status=status.HTTP_200_OK)
	# except:
	# 	return Response(status=status.HTTP_500_INTERNAL_SERVER_ERROR)

@api_view(["POST"])
def sign_pdf(request):
	try:
		list_id_pdf_file = request.data["list_id_pdf_file"]
		list_id=list_id_pdf_file.split(",")
		for i in list_id:
			pdf = pdfFile.objects.get(pk=int(i))
			pdf.confirmed = True
			pdf.save()
		print(list_id_pdf_file)
		return Response({"code":"00"}, status=status.HTTP_200_OK)
	except:
		err_mess = sys.exc_info()[0].__name__ + ": "+ str(sys.exc_info()[1])
		print(err_mess)
		return Response({"err":err_mess},status=status.HTTP_500_INTERNAL_SERVER_ERROR)

@api_view(["POST"])
def send_pdf(request):
	account_data={"user_name": "baynguyen2000@gmail.com",
  				"password": "Tu12345@"}
	headers = { 
	'Content-Type':'application/json' 
   }
	response_obj = requests.post(r"https://api-testing.pvs.com.vn/user-api/api/token/", data=json.dumps(account_data),headers=headers)
	token=response_obj.json()['data']['access']
	if response_obj.status_code >= 200 and response_obj.status_code<300:
		try:
			with transaction.atomic():
				list_id_pdf_file = request.data["list_id_pdf_file"]
				
				list_id=list_id_pdf_file.split(",")
				accountCate=[]
				for i in list_id:
					account = str(pdfFile.objects.get(pk=int(i)).account)
					if account not in accountCate:
						accountCate.append(account)
				for account in accountCate:
					listfile=[]
					for i in list_id:
						if account == str(pdfFile.objects.get(pk=int(i)).account):
							pdf=pdfFile.objects.get(pk=int(i))	
							pdffile=pdfFile.objects.get(pk=int(i)).slaveFile
							linkfile=settings.URL+"/"+str(pdffile)
							
							data_send ={
								"pdf_url":linkfile,
								"sign_pos": pdf.pos,
								"contact": "thach.nguyenphamngoc@kcc.com",
								"reason": "sign contract",
								"page_number":pdf.page_number
							}
							headers2 = { 'Content-Type':'application/json', 
							'Authorization': 'Bearer ' + token }

							print(data_send)
							response_obj2 = requests.post(r"https://api-testing.pvs.com.vn/e-invoice-api/api/ca-sign/sign-pdf/ 84", data=json.dumps(data_send), headers=headers2)
							binarytext=response_obj2.content
							
							with open("file.pdf","wb") as file:
								file.write(binarytext)
							with open ("file.pdf","rb") as file:
								name="ThưThôngBáo_{}_{}.pdf".format(account,str(datetime.now().date()))
								newpdf=pdf.slaveFile.save(name,File(file))
								pdf.sended=True
								pdf.sendingTime=datetime.now()
								pdf.save()
								fileurl= settings.URL+"/"+str(pdf.slaveFile)
						listfile.append(fileurl)
							
					for email in pdfFile.objects.get(pk=int(i)).emailExtracted.all():
						print(listfile)
						print(email)
						send_email.send_noti_to_partner_sign_by_email(listfile,str(email))
					# for email in 
					# system_pdf_link=settings.URL+"/"+str(pdf)
					# send_email.send_noti_to_partner_sign_by_email(system_pdf_link,"longnld@pvs.com.vn")
				return Response({"code":"00"}, status=status.HTTP_200_OK)
		except:
			err_mess = sys.exc_info()[0].__name__ + ": "+ str(sys.exc_info()[1])
			print(err_mess)
			return Response({"err":err_mess},status=status.HTTP_500_INTERNAL_SERVER_ERROR)
	else :
		return Response(response_obj.json()['message'], status=status.HTTP_500_INTERNAL_SERVER_ERROR)
@api_view(["POST"])
def deleteFile(request):
	print(request.data)
	pk=request.data["pdfFileID"]
	print(pk)
	if request.data["pdfFileID"]:
		fileWillBeDel=pdfFile.objects.get(pk=int(pk))
		print(fileWillBeDel)
		fileWillBeDel.delete()
	return Response({"msg":"delete success"}, status=status.HTTP_200_OK)

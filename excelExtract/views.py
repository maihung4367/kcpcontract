from django.shortcuts import render,redirect
import sys
from rest_framework.decorators import api_view
from excelExtract.forms import uploadDocumentForm
from excelExtract.models import document,pdfFile, excel
from . import excelExtract
from user.models import Profile
from django.db import transaction
from rest_framework import status
from django.conf import settings
from django.http import HttpResponse
from rest_framework.response import Response
from django.views.decorators.clickjacking import xframe_options_sameorigin
from utils import send_email
# import requests
import json
# Create your views here.
@xframe_options_sameorigin
def kcToolPage(request):
	form=uploadDocumentForm()
	files=document.objects.all()
	pdffiles=pdfFile.objects.all().order_by("-id")
	demoPdfFiles=pdfFile.objects.last()
	numberUnsignepdfs=len(pdfFile.objects.filter(signed=False))
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


def waitSignDoc(request):
	numberUnsignepdfs=len(pdfFile.objects.filter(signed=False))
	unsignedpdfs=pdfFile.objects.filter(signed=False).order_by("-id")
	return render(request,"KCtool/waitingsigndoc.html",{"numberUnsignepdfs":numberUnsignepdfs,"unsignedpdfs":unsignedpdfs, "URL":settings.URL, "active_id":2})

def signedDoc(request):
	numbersignepdfs=len(pdfFile.objects.filter(signed=True))
	pdfs=pdfFile.objects.filter(signed=True).order_by("sended")
	return render(request,"KCtool/signedDoc.html",{"numbersignepdfs":numbersignepdfs,"pdfs":pdfs, "active_id":3})

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
			annouce=excelExtract.exportFiles(loaict=values_category,fileID=id_excel,loaiAccount=values_account) 
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
			pdf.signed = True
			pdf.save()
		print(list_id_pdf_file)
		return Response({"code":"00"}, status=status.HTTP_200_OK)
	except:
		err_mess = sys.exc_info()[0].__name__ + ": "+ str(sys.exc_info()[1])
		print(err_mess)
		return Response({"err":err_mess},status=status.HTTP_500_INTERNAL_SERVER_ERROR)

@api_view(["POST"])
def send_pdf(request):
	
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
						pdf.sended=True
						pdf.save()
						fileurl= settings.URL+"/"+str(pdffile)
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
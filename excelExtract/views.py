from django.shortcuts import render,redirect
from rest_framework.decorators import api_view
from excelExtract.forms import uploadDocumentForm
from excelExtract.models import document,pdfFile, excel
from . import excelExtract
from django.db import transaction
from rest_framework import status
from django.conf import settings
from django.http import HttpResponse
from rest_framework.response import Response
from django.views.decorators.clickjacking import xframe_options_sameorigin
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
		print("12312312")
		with transaction.atomic():
			try:
				file=request.FILES['document']
				print(file)
				if file:
					excelExtract.importDataExcel(file)
			except:
				pass
	return render(request,"KCtool/KCtool.html",{"form":form,"files":files,"pdffiles":pdffiles,"demoPdfFiles":demoPdfFiles,"numberUnsignepdfs":numberUnsignepdfs})


def waitSignDoc(request):
	numberUnsignepdfs=len(pdfFile.objects.filter(signed=False))
	unsignedpdfs=pdfFile.objects.filter(signed=False).order_by("-id")
	return render(request,"KCtool/waitingsigndoc.html",{"numberUnsignepdfs":numberUnsignepdfs,"unsignedpdfs":unsignedpdfs, "URL":settings.URL})

def signedDoc(request):
	numberUnsignepdfs=len(pdfFile.objects.filter(signed=False))
	return render(request,"KCtool/signedDoc.html",{"numberUnsignepdfs":numberUnsignepdfs})
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
			excelExtract.exportFiles(loaict=values_category,fileID=id_excel,loaiAccount=values_account) 
			return Response({"code":"00"}, status=status.HTTP_200_OK)
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
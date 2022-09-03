from django.shortcuts import render,redirect
from rest_framework.decorators import api_view
from excelExtract.forms import uploadDocumentForm
from excelExtract.models import document,pdfFile
from . import excelExtract
from django.db import transaction
from rest_framework import status
from django.http import HttpResponse
from rest_framework.response import Response
from django.views.decorators.clickjacking import xframe_options_sameorigin
import requests
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
	return render(request,"KCtool/KCTool.html",{"form":form,"files":files,"pdffiles":pdffiles,"demoPdfFiles":demoPdfFiles,"numberUnsignepdfs":numberUnsignepdfs})


def waitSignDoc(request):
	numberUnsignepdfs=len(pdfFile.objects.filter(signed=False))
	unsignedpdfs=pdfFile.objects.filter(signed=False)
	return render(request,"KCtool/waitingsigndoc.html",{"numberUnsignepdfs":numberUnsignepdfs,"unsignedpdfs":unsignedpdfs})

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
			# # loaiCt= request.POST.getlist('loaiCt')
			# # loaiAccount =request.POST.getlist('loaiAccount')
			# loaiAccount=['acc1','acc2']
			# data_send = {
			# 		"master_file_id": id_excel,
			# 		# "is_publish_now": True,
					
					
			# 	}	
			# for account in loaiAccount:
			# 	if item["taxPercentage"] == 0:
			# 		item["taxPercentage"] = None
			# 	data_send["lines"].append({
			# 				"name": item["itemName"],
			# 				"code": "1",
			# 				"unit": item["unitName"],
			# 				"quantity": int(item["quantity"]),
			# 				"price": int(item["unitPrice"]),
			# 				"type": "HANG_HOA",
			# 				"rate": item["taxPercentage"]
			# 				})
			# print(data_send)
			# api_url = settings.E_INVOICE_API_URL +"/e-invoice-service-api/api/e-invoices/replace/{}".format(int(invoice.id_invoice_eInvoice))

			# response_obj = requests.post(api_url, data=json.dumps(data_send), headers=headers)
			# print(response_obj.json())

			# if response_obj.status_code >= 200 and response_obj.status_code<300:
				# return Response(response_obj.json(), status=status.HTTP_200_OK)
			# return Response({"code":"00"}, status=status.HTTP_200_OK)
			# else:
			# 	return Response(response_obj.json()['message'], status=status.HTTP_500_INTERNAL_SERVER_ERROR)

			excelExtract.exportFiles(loaict='ALL',fileID=id_excel,loaiAccount='ALL') 
			return Response({"code":"00"}, status=status.HTTP_200_OK)
	except:
		return Response(status=status.HTTP_500_INTERNAL_SERVER_ERROR)



	
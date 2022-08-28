from django.shortcuts import render,redirect
from rest_framework.decorators import api_view
from excelExtract.forms import uploadDocumentForm
from excelExtract.models import document,pdfFile
from . import excelExtract
from django.http import HttpResponse
from rest_framework.response import Response
from django.views.decorators.clickjacking import xframe_options_sameorigin
# Create your views here.
@xframe_options_sameorigin
def kcToolPage(request):
    form=uploadDocumentForm()
    files=document.objects.all()
    pdffiles=pdfFile.objects.all().order_by("-id")
    demoPdfFiles=pdfFile.objects.first()
    if request.method=='POST':
        form=uploadDocumentForm(request.POST,request.FILES)
        if form.is_valid():
            file=request.FILES.get('document')
            excelExtract.importDataExcel(file)
    return render(request,"KCtool/KCTool.html",{"form":form,"files":files,"pdffiles":pdffiles,"demoPdfFiles":demoPdfFiles})
# def excelToListPdfs(request):  
#             return Response(request,"KCtool/KCTool.html")
@api_view(["POST"])
def getIdList(request):
    if request.method=='POST':
        print(request.data)
        loaict=request.data['loaict']
        listId=request.POST.getlist('file')
        for  f in listId:
            loaiAccount=request.data['fileID{}'.format(f)]
            excelExtract.exportFiles(loaict=loaict,fileID=f,loaiAccount=loaiAccount)  
        return redirect("KCTool:kcToolPage")
    
    
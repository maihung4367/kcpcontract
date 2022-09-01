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
    demoPdfFiles=pdfFile.objects.last()
    numberUnsignepdfs=len(pdfFile.objects.filter(signed=False))
    if request.method=='POST':
        print("12312312")
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
    
    
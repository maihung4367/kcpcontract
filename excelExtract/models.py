from django.db import models
from user.models import Profile
import os
def fileExtensionValidate(value):
    from django.core.exceptions import ValidationError
    ext = os.path.splitext(value.name)[-1]  # [0] returns path+filename
    valid_extensions = ['.xlsx',]
    if not ext.lower() in valid_extensions:
        raise ValidationError('Unsupported file extension.')
class document(models.Model):
    upload_by = models.ForeignKey(Profile, on_delete=models.SET_NULL, null=True, blank=True)
    uploadTime=models.DateTimeField(auto_now_add=True,blank=True,null=True)
    document = models.FileField(upload_to="documents",validators=[fileExtensionValidate])
    def __str__(self):
        return "{}".format(str(self.document))
class excelAccount(models.Model):
    account=models.CharField(max_length=30,unique=True)  
    def __str__(self):
        return "{}".format(str(self.account))      
class accountEmail(models.Model):
    account=models.ForeignKey(excelAccount,on_delete=models.CASCADE)
    email=models.EmailField(blank=True,null=True)
    def __str__(self):
        return "{}".format(str(self.email)) 
class excel(models.Model):
    filename                =models.ForeignKey(document, on_delete=models.CASCADE)
    #group,account,postStartDate,postEndDate,mechanicsGetORDiscount,noiDungChuongTrinh,budgetRir,loaiCt
    group                   =models.CharField(max_length=12)
    account                 =models.CharField(max_length=30)
    postStartDate           =models.DateField(null=True,blank=True)
    postEndDate             =models.DateField(null=True,blank=True)
    mechanicsGetORDiscount  =models.TextField(null=True,blank=True)
    noiDungChuongTrinh      =models.TextField(null=True,blank=True)
    product                 =models.TextField(null=True,blank=True)
    budgetRir               =models.CharField(max_length=20,null=True,blank=True)
    loaiCt                  =models.CharField(max_length=20,null=True,blank=True)

    def __str__(self):
        return "{}".format(str(self.filename))
class pdfFile(models.Model):
    masterFile=models.ForeignKey(document, on_delete=models.CASCADE)
    slaveFile=models.FileField(upload_to="documents/slavefiles")
    account=models.ForeignKey(excelAccount,on_delete=models.CASCADE)
    loaict=models.TextField(blank=True,null=True)
    createdTime=models.DateTimeField(auto_now_add=True,blank=True,null=True)
    sendingTime=models.DateTimeField(blank=True,null=True)
    emailExtracted=models.ManyToManyField(accountEmail,blank=True)
    signed=models.BooleanField(default=False)   
    sended=models.BooleanField(default=False) 
    def __str__(self):
        return "{}".format(str(self.slaveFile).replace("documents/slavefiles/",""))
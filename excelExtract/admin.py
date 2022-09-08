from django.contrib import admin
from.models import document,excel,pdfFile,excelAccount,accountEmail
# Register your models here.
admin.site.register(document)
admin.site.register(excel)
admin.site.register(pdfFile)
admin.site.register(excelAccount)
admin.site.register(accountEmail)
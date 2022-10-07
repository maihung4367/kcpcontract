from django.contrib import admin
from.models import document,excel,pdfFile,excelAccount,accountEmail
# Register your models here.
@admin.action(description='updateAllStandardName')
def updateAllStandardName(self, request, queryset):
    for object in queryset:
        object.standardName=str(object.account).replace(" ","").lower()
        print(str(object.account).replace(" ","").lower())
        object.save()
class ExcelAcountAdmin(admin.ModelAdmin):
    model =excelAccount
    actions = [updateAllStandardName] 

    
admin.site.register(document)
admin.site.register(excel)
admin.site.register(pdfFile)
admin.site.register(excelAccount,ExcelAcountAdmin)
admin.site.register(accountEmail)
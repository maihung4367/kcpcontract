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
    list_display = ('account','responsibleBy',)
class AccountEmailAdmin(admin.ModelAdmin):
    list_display = ('email','account',)  # Add the fields you want to display here
class pdfFileAdmin(admin.ModelAdmin):
    list_display = ('__str__','account','loaict','confirmed','signed','sended','is_deleted')  # Add the fields you want to display here    
admin.site.register(document)
admin.site.register(excel)
admin.site.register(pdfFile,pdfFileAdmin)
admin.site.register(excelAccount,ExcelAcountAdmin)
admin.site.register(accountEmail,AccountEmailAdmin)
from django.urls import path,include
from . import views
app_name='KCTool'
urlpatterns = [
    path('',views.kcToolPage,name="kcToolPage" ),
    path("getIdList",views.getIdList,name="getIdList"),
    path("waitingConfirmDocs/",views.waitConfirmDoc,name="waitConfirmDoc"),
    path("signedDocs/",views.signedDoc,name="signedDoc"),
    path("create-pdf", views.create_pdf, name="create_pdf"),
    path("getListAccount", views.getListAccount, name="getListAccount"),
    path("sign_pdf", views.sign_pdf, name="sign_pdf"),
    path("send_pdf",views.send_pdf, name="send_pdf"),
    path("delete_file",views.deleteFile, name="delete_file"),
    path("delete_excel_file",views.deleteExcelFile, name="delete_file_excel"),
    path('update-profile', views.update_profile, name="update_profile")


]

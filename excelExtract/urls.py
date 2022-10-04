from django.urls import path,include
from . import views
app_name='KCTool'
urlpatterns = [
    path('',views.kcToolPage,name="kcToolPage" ),
    path("NewCreatedDocs/",views.newCreatedDocs,name="newCreatedDocs"),
    path("ConfirmedDocs/",views.confirmedDocs,name="confirmedDocs"),
    path("SignedDocs/",views.signedDocs,name="signedDocs"),
    path("create-pdf", views.create_pdf, name="create_pdf"),
    path("getListAccount", views.getListAccount, name="getListAccount"),
    path("confirm_pdf", views.confirm_pdf, name="confirm_pdf"),
    path("sign_and_send_pdf",views.sign_and_send_pdf, name="sign_and_send_pdf"),
    path("send_pdf",views.send_pdf, name="send_pdf"),
    path("delete_file",views.deleteFile, name="delete_file"),
    path("delete_excel_file",views.deleteExcelFile, name="delete_file_excel"),
    path('update-profile', views.update_profile, name="update_profile"),
    path("UntrackedDocs/",views.untrackedDocs,name="untrackedDocs")


]

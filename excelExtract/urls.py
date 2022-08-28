from django.urls import path,include
from . import views
app_name='KCTool'
urlpatterns = [
    path('',views.kcToolPage,name="kcToolPage" ),
    path("getIdList",views.getIdList,name="getIdList")
]

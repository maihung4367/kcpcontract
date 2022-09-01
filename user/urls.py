from django.urls import path,include
from . import views
app_name='KCTool'
urlpatterns = [
  path('', views.loginView, name='login')
]

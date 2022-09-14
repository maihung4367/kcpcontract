from django.urls import path,include
from . import views
from django.contrib.auth import views as auth_views
app_name='KCTool'
urlpatterns = [
  path('', views.loginView, name='login'),
  path('^logout/$',auth_views.LogoutView.as_view(),name='logout'),
]

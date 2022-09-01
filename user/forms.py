from django.contrib.auth.forms import UserCreationForm, UserChangeForm

from django import forms
from .models import User






class CustomUserCreationForm(UserCreationForm):

    class Meta(UserCreationForm):
        model = User
        fields = ('user_name','is_staff','is_active','is_admin')

class CustomUserChangeForm(UserChangeForm):

    class Meta:
        model = User
        fields = ('user_name','is_staff','is_active','is_admin','password',)

class LoginForm(forms.Form):
	username = forms.CharField(max_length=200, widget=forms.TextInput(attrs={'class': 'input-user', 'placeholder':"Tên đăng nhập", 'style':'font-size: 16px'}))
    
	password = forms.CharField(max_length=200, widget=forms.PasswordInput(attrs={'class': 'input-password', 'placeholder':"Mật Khẩu", 'id':"password-field", 'name':"password", 'style':'font-size: 16px'}))
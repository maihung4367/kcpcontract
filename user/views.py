from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login
from .forms import LoginForm
from django.contrib.auth.decorators import login_required
from datetime import datetime
import logging

login_tracker=logging.getLogger("login_track")
# Create your views here.


def loginView(request):
	form = LoginForm()
	if request.user.is_authenticated:
		return redirect('KCTool:kcToolPage')
	if request.method == 'POST':
		username = request.POST.get('username')
		password = request.POST.get('password')
		user = authenticate(request, username=username, password=password)
		if user :
			login(request, user)
			login_tracker.info("SUCCESS--"+str(username)+"//"+str(datetime.now()))
			return redirect('KCTool:kcToolPage')
		else:
			login_tracker.info("DEBUG_TRACK--"+"Authen:"+str(username)+"//"+str(password)+"--Cookies:"+str(request.COOKIES))
			
			return render(request, 'login.html', {'form':form,"msg":"Please check the password/account"})
	return render(request, 'login.html', {'form':form})
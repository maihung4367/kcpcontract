from django.shortcuts import render, redirect
from django.contrib.auth import authenticate, login
from .forms import LoginForm
from django.contrib.auth.decorators import login_required

# Create your views here.


def loginView(request):
	if request.method == 'POST':
		username = request.POST.get('username')
		password = request.POST.get('password')
		user = authenticate(request, username=username, password=password)
		if user is not None:
			login(request, user)
			return redirect('KCTool:kcToolPage')
		else:
			form = LoginForm()
			return render(request, 'login.html', {'form':form})
	form = LoginForm()
	return render(request, 'login.html', {'form':form})
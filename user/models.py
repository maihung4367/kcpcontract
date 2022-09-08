from django.db import models
from .utils_models import UserManager
from django.conf import settings
from django.contrib.auth.models import AbstractBaseUser,PermissionsMixin
# Create your models here.


class Profile(models.Model):
    full_name = models.CharField(max_length=255, null=True, blank = True)
    user = models.OneToOneField(settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.CASCADE)
    phone_number = models.CharField(max_length=15, null=True, blank = True)
    position = models.CharField(max_length=255, null=True, blank = True)
    email = models.EmailField(null=True, blank = True)
    address = models.CharField(max_length=255, null=True, blank = True)
    company_name = models.CharField(max_length=255, null=True, blank = True)


class User(AbstractBaseUser, PermissionsMixin):
    user_name = models.CharField("user name",max_length=254,unique=True)
    email = models.EmailField("email address", max_length=254,unique=False)
    is_staff = models.BooleanField("staff status", default=False)
    is_active = models.BooleanField(
        "active", default=True)
    is_admin = models.BooleanField(
        "admin PVS", default=False)

    USERNAME_FIELD = "user_name"

    objects = UserManager()

    def __str__(self):        
        return self.user_name
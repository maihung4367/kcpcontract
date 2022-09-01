from django.db import models
from .utils_models import UserManager
from django.contrib.auth.models import AbstractBaseUser,PermissionsMixin
# Create your models here.


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
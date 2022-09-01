from django.contrib.auth.models import BaseUserManager
from django.contrib.auth import  get_user_model

class UserManager(BaseUserManager):
    use_in_migrations = True

    def _create_user(
            self, user_name, password, **kwargs):
        #email = self.normalize_email(email)
        is_staff = kwargs.pop('is_staff', False)
        is_superuser = kwargs.pop(
            'is_superuser', False)
        user = self.model(
            user_name=user_name,
            is_active=True,
            is_staff=is_staff,
            is_superuser=is_superuser,
            **kwargs)
        user.set_password(password)
        user.save(using=self._db)
        return user

    def create_user(
            self, user_name, password=None,
            **extra_fields):
        return self._create_user(
            user_name, password, **extra_fields)

    def create_superuser(
            self, user_name, password,
            **extra_fields):
        return self._create_user(
            user_name, password,
            is_staff=True, is_superuser=True,
            **extra_fields)

    def update_pw_user(self, user_name, password):
        #user_name = self.normalize_email(user_name)
        user = get_user_model().objects.get(user_name=user_name)
        user.set_password(password)
        user.save()
        return user
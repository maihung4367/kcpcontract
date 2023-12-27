from django.contrib import admin
from .models import  User, Profile
from .forms import CustomUserChangeForm, CustomUserCreationForm
from django.contrib.auth.admin import UserAdmin
# Register your models here.

class ProfileAdmin(admin.ModelAdmin):
    list_display = ('full_name', 'email')

class CustomUserAdmin(UserAdmin):
	
    add_form = CustomUserCreationForm
    form = CustomUserChangeForm
    list_display = ["user_name", "is_staff",
    "is_active",'is_signer','is_uploader', "pk", "last_login"]
    model = User
    ordering = ["user_name", "last_login"]
    fieldsets = ((None, {'fields': ('user_name',)}), ('Password',{'fields': ('password',)}), ('dates',{'fields':('last_login',)}), ('Permissions', {'fields': ('is_admin','is_staff','is_active','is_superuser','is_signer','is_uploader','user_permissions','groups')}),)

    add_fieldsets = ((None,{
        'classes': ('wide',),
        'fields': ('user_name', 'password1', 'password2', 'is_admin','is_staff','is_active','is_signer','is_uploader','is_superuser','user_permissions','groups')}),
    )

    search_fields = ('user_name',"pk","groups__name")

    def group(self,obj):
        groups = []
        for group in obj.groups.all():
            groups.append(group.name)
        return " ".join(groups)
    group.short_description = "Groups"

admin.site.register(User, CustomUserAdmin)
admin.site.register(Profile,ProfileAdmin)

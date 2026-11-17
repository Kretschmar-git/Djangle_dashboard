from django.contrib import admin
from django.contrib.auth.admin import UserAdmin

from .models import CustomUser, DashboardMetric


@admin.register(CustomUser)
class CustomUserAdmin(UserAdmin):
    model = CustomUser
    list_display = ["username", "email", "role", "is_staff"]
    fieldsets = UserAdmin.fieldsets + (("Perfil", {"fields": ("role",)}),)
    add_fieldsets = UserAdmin.add_fieldsets + (("Perfil", {"fields": ("role",)}),)


@admin.register(DashboardMetric)
class DashboardMetricAdmin(admin.ModelAdmin):
    list_display = ["title", "date", "result_value", "status", "sort_order"]
    search_fields = ["title", "target_text"]
    list_filter = ["date", "status"]

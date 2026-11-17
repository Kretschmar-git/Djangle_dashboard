from django.contrib.auth.models import AbstractUser
from django.db import models
from django.utils import timezone


class CustomUser(AbstractUser):
    ROLE_ADMIN = "admin"
    ROLE_VIEWER = "viewer"
    ROLE_CHOICES = [
        (ROLE_ADMIN, "Administrador"),
        (ROLE_VIEWER, "Visualizador"),
    ]

    role = models.CharField(max_length=20, choices=ROLE_CHOICES, default=ROLE_VIEWER)

    def __str__(self):
        return f"{self.username} ({self.get_role_display()})"


class ActiveCompany(models.Model):
    name = models.CharField(max_length=200, unique=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["name"]

    def __str__(self):
        return self.name


class MeetingParticipant(models.Model):
    name = models.CharField(max_length=200, unique=True)
    role = models.CharField(max_length=100, blank=True, default="")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["name"]

    def __str__(self):
        return self.name


class Meeting(models.Model):
    title = models.CharField(max_length=200)
    date = models.DateField(default=timezone.now)
    companies = models.ManyToManyField(ActiveCompany, blank=True, related_name="meetings")
    participants = models.ManyToManyField(MeetingParticipant, blank=True, related_name="meetings")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-date", "title"]

    def __str__(self):
        return self.title


class DashboardMetric(models.Model):
    STATUS_ACHIEVED = "achieved"
    STATUS_EXCEEDED = "exceeded"
    STATUS_NOT_ACHIEVED = "not_achieved"
    STATUS_WARNING = "warning"
    STATUS_NA = "not_applicable"

    STATUS_CHOICES = [
        (STATUS_ACHIEVED, "Meta atendida"),
        (STATUS_EXCEEDED, "Meta superada"),
        (STATUS_NOT_ACHIEVED, "Meta não atendida"),
        (STATUS_WARNING, "Sem meta para o 1º ano"),
        (STATUS_NA, "Não aplicável"),
    ]

    date = models.DateField(default=timezone.now)
    title = models.CharField(max_length=200)
    target_text = models.CharField(max_length=300)
    result_value = models.CharField(max_length=100, blank=True, default="")
    status = models.CharField(max_length=30, choices=STATUS_CHOICES, default=STATUS_ACHIEVED)
    sort_order = models.PositiveIntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["sort_order", "id"]

    @property
    def status_label(self):
        labels = {
            self.STATUS_ACHIEVED: "Meta atendida",
            self.STATUS_EXCEEDED: "Meta superada",
            self.STATUS_NOT_ACHIEVED: "Meta não atendida",
            self.STATUS_WARNING: "Sem meta para o 1º ano",
            self.STATUS_NA: "Não aplicável",
        }
        return labels.get(self.status, "Meta atendida")

    def __str__(self):
        return f"{self.title} - {self.result_value}"

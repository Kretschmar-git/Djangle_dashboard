from datetime import date
from functools import wraps

from django.contrib import messages
from django.contrib.auth import logout
from django.contrib.auth.decorators import login_required
from django.contrib.auth.views import LoginView
from django.core.exceptions import PermissionDenied
from django.db.models import Max
from django.http import HttpResponse
from django.shortcuts import redirect, render
from django.urls import reverse_lazy
from openpyxl import Workbook

from .forms import ExcelUploadForm
from .models import ActiveCompany, CustomUser, DashboardMetric, Meeting
from .services import import_dashboard_excel


def home(request):
    if request.user.is_authenticated:
        return redirect("dashboard")
    return redirect("login")


class CustomLoginView(LoginView):
    template_name = "dashboard/login.html"
    redirect_authenticated_user = True

    def get_success_url(self):
        return reverse_lazy("dashboard")


def logout_view(request):
    logout(request)
    return redirect("login")


def admin_required(view_func):
    @wraps(view_func)
    def _wrapped(request, *args, **kwargs):
        if not request.user.is_authenticated:
            return redirect("login")
        if request.user.role != CustomUser.ROLE_ADMIN:
            raise PermissionDenied("Acesso restrito ao administrador.")
        return view_func(request, *args, **kwargs)

    return _wrapped


def parse_metric_result_value(value):
    if value is None:
        return 0, "", "NA", "0"

    text = str(value).strip()
    if not text or text.upper() == "NA":
        return 0, "", "NA", "0"

    suffix = "%" if text.endswith("%") else ""
    numeric_text = text.replace("%", "").replace(".", "").replace(",", ".")

    try:
        parsed = float(numeric_text)
    except ValueError:
        return 0, "", text, "0"

    normalized = format(parsed, "f").rstrip("0").rstrip(".") if "." in format(parsed, "f") else str(int(parsed))
    return parsed, suffix, text, normalized


def build_analytics_chart_series(metrics):
    status_colors = {
        DashboardMetric.STATUS_ACHIEVED: "#1e9d6c",
        DashboardMetric.STATUS_EXCEEDED: "#3b82f6",
        DashboardMetric.STATUS_NOT_ACHIEVED: "#dc4c4c",
        DashboardMetric.STATUS_WARNING: "#f4a03d",
        DashboardMetric.STATUS_NA: "#c7d0ca",
    }

    chart_series = []
    for metric in metrics:
        value, suffix, display_value, value_raw = parse_metric_result_value(metric.result_value)
        chart_series.append(
            {
                "label": metric.title,
                "value": value,
                "value_raw": value_raw,
                "display_value": display_value,
                "suffix": suffix,
                "color": status_colors.get(metric.status, "#1e9d6c"),
                "status_label": metric.status_label,
            }
        )
    return chart_series


def apply_metric_filters(request, queryset):
    selected_year = request.GET.get("year")
    selected_month = request.GET.get("month")
    selected_status = request.GET.get("status")
    search_term = request.GET.get("search", "").strip()
    date_from = request.GET.get("date_from")
    date_to = request.GET.get("date_to")

    if selected_year:
        queryset = queryset.filter(date__year=selected_year)
    if selected_month:
        queryset = queryset.filter(date__month=selected_month)
    if selected_status:
        queryset = queryset.filter(status=selected_status)
    if search_term:
        queryset = queryset.filter(title__icontains=search_term)
    if date_from:
        queryset = queryset.filter(date__gte=date_from)
    if date_to:
        queryset = queryset.filter(date__lte=date_to)

    return queryset, {
        "selected_year": selected_year,
        "selected_month": selected_month,
        "selected_status": selected_status,
        "search_term": search_term,
        "date_from": date_from,
        "date_to": date_to,
    }


@login_required
def dashboard_view(request):
    metrics = DashboardMetric.objects.all().order_by("sort_order", "id")
    metrics, filters = apply_metric_filters(request, metrics)
    latest_metric_ids = (
        metrics.order_by()
        .values("title")
        .annotate(latest_id=Max("id"))
        .values("latest_id")
    )
    metrics = DashboardMetric.objects.filter(id__in=latest_metric_ids).order_by("sort_order", "id")

    years = DashboardMetric.objects.dates("date", "year", order="DESC")
    months = DashboardMetric.objects.dates("date", "month", order="DESC")
    status_choices = [choice for choice in DashboardMetric.STATUS_CHOICES]
    active_companies = ActiveCompany.objects.filter(is_active=True).order_by("name")
    meetings = Meeting.objects.prefetch_related("participants", "companies").order_by("-date")[:5]

    return render(
        request,
        "dashboard/dashboard.html",
        {
            "metrics": metrics,
            "years": years,
            "months": months,
            "status_choices": status_choices,
            "user": request.user,
            "active_companies": active_companies,
            "meetings": meetings,
            **filters,
        },
    )


@login_required
def analytics_view(request):
    metrics = DashboardMetric.objects.all().order_by("sort_order", "id")
    metrics, filters = apply_metric_filters(request, metrics)
    chart_series = build_analytics_chart_series(metrics)

    labels = [item["label"] for item in chart_series]
    values = [item["value"] for item in chart_series]

    achieved = metrics.filter(status=DashboardMetric.STATUS_ACHIEVED).count()
    exceeded = metrics.filter(status=DashboardMetric.STATUS_EXCEEDED).count()
    not_achieved = metrics.filter(status=DashboardMetric.STATUS_NOT_ACHIEVED).count()
    warning = metrics.filter(status=DashboardMetric.STATUS_WARNING).count()
    na_count = metrics.filter(status=DashboardMetric.STATUS_NA).count()

    return render(
        request,
        "dashboard/analytics.html",
        {
            "metrics": metrics,
            "labels": labels,
            "values": values,
            "chart_series": chart_series,
            "achieved": achieved,
            "exceeded": exceeded,
            "not_achieved": not_achieved,
            "warning": warning,
            "na_count": na_count,
            "years": DashboardMetric.objects.dates("date", "year", order="DESC"),
            "months": DashboardMetric.objects.dates("date", "month", order="DESC"),
            "status_choices": DashboardMetric.STATUS_CHOICES,
            "user": request.user,
            **filters,
        },
    )


@login_required
@admin_required
def admin_import_view(request):
    if request.method == "POST":
        form = ExcelUploadForm(request.POST, request.FILES)
        if form.is_valid():
            try:
                imported_rows = import_dashboard_excel(form.cleaned_data["file"])
                messages.success(
                    request,
                    f"{imported_rows} novos indicadores registrados. "
                    "Os dados existentes foram mantidos.",
                )
                return redirect("dashboard")
            except Exception as exc:  # pragma: no cover - comportamento do usuário
                messages.error(request, f"Não foi possível importar a planilha: {exc}")
    else:
        form = ExcelUploadForm()

    return render(request, "dashboard/admin_import.html", {"form": form, "metrics": DashboardMetric.objects.all()})


def download_template(request):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Dashboard"
    worksheet.append(["data", "indicador", "meta", "resultado", "status", "ordem"])
    worksheet.append(["2026-01-31", "Empresas ativas", "Manter ou aumentar a quantidade atual", "14", "Meta atendida", "1"])
    worksheet.append(["2026-01-31", "Presença dos nucleados", "Mínimo de 70% dos nucleados ativos", "82,69%", "Meta atendida", "2"])
    worksheet.append(["2026-01-31", "Reuniões realizadas", "Cumprir o calendário definido pelo núcleo", "100%", "Meta atendida", "3"])
    worksheet.append(["2026-01-31", "Novas empresas nucleadas", "Ter pelo menos 3 novas empresas nucleadas por ano", "6", "Meta superada", "4"])
    worksheet.append(["2026-01-31", "PANs realizados no ano", "Meta anual definida por cada núcleo no planejamento", "0", "Sem meta para o 1º ano", "5"])
    worksheet.append(["2026-01-31", "Execução do plano de ação (ações, eventos, capacitações)", "Mínimo de 80% conforme planejamento estratégico", "92%", "Meta atendida", "6"])
    worksheet.append(["2026-01-31", "Satisfação dos nucleados", "Nota mínima 8,0", "NA", "Não aplicável", "7"])

    response = HttpResponse(content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response["Content-Disposition"] = 'attachment; filename="modelo_dashboard.xlsx"'
    workbook.save(response)
    return response

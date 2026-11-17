from datetime import datetime
from io import BytesIO

from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase
from openpyxl import Workbook

from .models import ActiveCompany, CustomUser, DashboardMetric, Meeting, MeetingParticipant
from .services import import_dashboard_excel
from .views import build_analytics_chart_series


class DashboardModelTests(TestCase):
    def test_user_roles_are_defined(self):
        self.assertIn(CustomUser.ROLE_ADMIN, ["admin", "viewer"])
        self.assertIn(CustomUser.ROLE_VIEWER, ["admin", "viewer"])

    def test_dashboard_metric_can_be_created(self):
        metric = DashboardMetric.objects.create(
            title="Empresas ativas",
            target_text="Manter ou aumentar a quantidade atual",
            result_value="14",
            status="achieved",
            sort_order=1,
        )
        self.assertEqual(metric.title, "Empresas ativas")
        self.assertEqual(metric.status, "achieved")

    def test_analytics_chart_series_keeps_actual_metric_values(self):
        DashboardMetric.objects.create(
            title="Empresas ativas",
            target_text="Meta",
            result_value="14",
            status="achieved",
            sort_order=1,
        )
        DashboardMetric.objects.create(
            title="Presença dos nucleados",
            target_text="Meta",
            result_value="82,69%",
            status="achieved",
            sort_order=2,
        )

        chart_series = build_analytics_chart_series(DashboardMetric.objects.all())

        self.assertEqual(chart_series[0]["value"], 14)
        self.assertEqual(chart_series[1]["value"], 82.69)
        self.assertEqual(chart_series[1]["display_value"], "82,69%")

    def test_dashboard_lists_active_companies_and_meeting_participants(self):
        user = CustomUser.objects.create_user(username="admin", password="senha123", role=CustomUser.ROLE_ADMIN)
        company = ActiveCompany.objects.create(name="Metal Bogo")
        participant = MeetingParticipant.objects.create(name="Andressa B. Vick")
        meeting = Meeting.objects.create(title="Reunião de acompanhamento")
        meeting.companies.add(company)
        meeting.participants.add(participant)

        self.client.force_login(user)
        response = self.client.get("/dashboard/")

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Metal Bogo")
        self.assertContains(response, "Andressa B. Vick")
        rendered_page = response.content.decode()
        self.assertLess(rendered_page.index('class="table-panel"'), rendered_page.index('class="summary-grid"'))

    def test_dashboard_shows_only_latest_version_of_each_metric(self):
        user = CustomUser.objects.create_user(username="admin", password="senha123", role=CustomUser.ROLE_ADMIN)
        DashboardMetric.objects.create(
            title="Empresas ativas",
            target_text="Meta",
            result_value="14",
            status=DashboardMetric.STATUS_ACHIEVED,
            sort_order=1,
        )
        DashboardMetric.objects.create(
            title="Empresas ativas",
            target_text="Meta",
            result_value="17",
            status=DashboardMetric.STATUS_ACHIEVED,
            sort_order=1,
        )

        self.client.force_login(user)
        response = self.client.get("/dashboard/")

        self.assertEqual(response.status_code, 200)
        displayed_metrics = list(response.context["metrics"])
        self.assertEqual(len(displayed_metrics), 1)
        self.assertEqual(displayed_metrics[0].result_value, "17")

    def test_dashboard_does_not_recreate_example_data(self):
        user = CustomUser.objects.create_user(username="admin", password="senha123", role=CustomUser.ROLE_ADMIN)
        self.client.force_login(user)

        response = self.client.get("/dashboard/")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(DashboardMetric.objects.count(), 0)
        self.assertEqual(ActiveCompany.objects.count(), 0)
        self.assertEqual(Meeting.objects.count(), 0)
        self.assertEqual(MeetingParticipant.objects.count(), 0)

    def test_imports_multi_sheet_nucleus_workbook(self):
        workbook = Workbook()
        settings = workbook.active
        settings.title = "INDICADORES"
        settings.append(["CONFIGURAÇÕES"])
        settings.append(["Data da última atualização", datetime(2026, 10, 28)])

        report = workbook.create_sheet("RELATÓRIO TRIMESTRAL")
        report.append(["RELATÓRIO TRIMESTRAL — MODELO"])
        report.append(["Instruções"])
        report.append([])
        report.append(["Indicador", "Resultado atual", "Meta / referência", "Situação", "Fonte"])
        report.append(["Associação", "ACIAM", None, None, None])
        report.append(["Núcleo", "Núcleo de Teste", None, None, None])
        report.append(["Ano", 2026, None, None, None])
        report.append(["Presença", 0.99, 0.7, "Meta inicial 70%", "PRESENÇAS"])
        report.append(["Novas empresas", 4, 3, "Meta inicial 3/ano", "EMPRESAS"])
        report.append(["Satisfação", None, 8, "Meta inicial 8,0", "SATISFAÇÃO"])

        companies = workbook.create_sheet("EMPRESAS")
        companies.append(["CADASTRO DE EMPRESAS"])
        companies.append(["Revisar dados"])
        companies.append(["ID empresa", "Empresa", "Associação", "Núcleo", "Entrada", "Saída", "Situação"])
        companies.append(["E001", "Empresa Teste", "ACIAM", "Núcleo", None, None, "Ativa"])

        participants = workbook.create_sheet("PARTICIPANTES")
        participants.append(["CADASTRO DE PARTICIPANTES"])
        participants.append(["Revisar dados"])
        participants.append(["ID participante", "Nome", "Empresa", "Cargo"])
        participants.append(["P001", "Pessoa Teste", "Empresa Teste", "Analista"])

        meetings = workbook.create_sheet("REUNIÕES")
        meetings.append(["REGISTRO DE REUNIÕES"])
        meetings.append(["Revisar dados"])
        meetings.append([])
        meetings.append(["ID reunião", "Data", "Associação", "Núcleo", "Situação", "Planejada?", "Tema/observações"])
        meetings.append(["R001", datetime(2026, 10, 21), "ACIAM", "Núcleo", "Realizada", "Sim", "Planejamento"])

        attendance = workbook.create_sheet("PRESENÇAS")
        attendance.append(["PRESENÇAS"])
        attendance.append(["Revisar dados"])
        attendance.append(
            ["ID registro", "ID reunião", "Data da reunião", "ID participante", "Participante", "Empresa", "Status presença"]
        )
        attendance.append(["PR001", "R001", datetime(2026, 10, 21), "P001", "Pessoa Teste", "Empresa Teste", "Presente"])
        attendance.append(["PR002", "R001", datetime(2026, 10, 21), "P002", "Pessoa Ausente", "Empresa Teste", "Ausente"])
        attendance.append(["PR003", "R001", datetime(2026, 10, 21), "P003", "Pessoa Justificada", "Empresa Teste", "Justificada"])
        attendance.append(["PR004", "R001", datetime(2026, 10, 21), "P004", "Pessoa N/A", "Empresa Teste", "N/A"])

        payload = BytesIO()
        workbook.save(payload)
        upload = SimpleUploadedFile(
            "nucleo.xlsx",
            payload.getvalue(),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        imported_count = import_dashboard_excel(upload)

        self.assertEqual(imported_count, 3)
        presence = DashboardMetric.objects.get(title="Presença")
        self.assertEqual(presence.result_value, "33,33%")
        self.assertEqual(presence.target_text, "70%")
        self.assertEqual(presence.status, DashboardMetric.STATUS_NOT_ACHIEVED)
        self.assertEqual(DashboardMetric.objects.get(title="Novas empresas").status, DashboardMetric.STATUS_EXCEEDED)
        self.assertEqual(DashboardMetric.objects.get(title="Satisfação").status, DashboardMetric.STATUS_NA)
        self.assertEqual(presence.date.isoformat(), "2026-10-28")

        meeting = Meeting.objects.get()
        self.assertEqual(meeting.date.isoformat(), "2026-10-21")
        self.assertEqual(meeting.participants.get().name, "Pessoa Teste")
        self.assertEqual(meeting.participants.get().role, "Analista")
        self.assertEqual(meeting.companies.get().name, "Empresa Teste")

    def test_import_keeps_unlisted_companies_and_ignores_presence_only_companies(self):
        ActiveCompany.objects.create(name="Empresa Legada")

        workbook = Workbook()
        report = workbook.active
        report.title = "RELATÓRIO TRIMESTRAL"
        report.append(["Indicador", "Resultado atual", "Meta / referência", "Situação"])
        report.append(["Empresas ativas", 1, "Manter", "Validar cadastro"])

        companies = workbook.create_sheet("EMPRESAS")
        companies.append(["ID empresa", "Empresa", "Situação"])
        companies.append(["E001", "Empresa Ativa", "Ativa"])
        companies.append(["E002", "Empresa Inativa", "Inativa"])

        attendance = workbook.create_sheet("PRESENÇAS")
        attendance.append(
            ["ID registro", "ID reunião", "Data da reunião", "ID participante", "Participante", "Empresa", "Status presença"]
        )
        attendance.append(["PR001", "R001", datetime(2026, 10, 21), "P001", "Pessoa 1", "Empresa Ativa", "Presente"])
        attendance.append(["PR002", "R001", datetime(2026, 10, 21), "P002", "Pessoa 2", "Empresa Inativa", "Presente"])
        attendance.append(["PR003", "R001", datetime(2026, 10, 21), "P003", "Pessoa 3", "Empresa fora do cadastro", "Presente"])

        payload = BytesIO()
        workbook.save(payload)
        upload = SimpleUploadedFile("nucleo.xlsx", payload.getvalue())

        import_dashboard_excel(upload)

        active_names = list(
            ActiveCompany.objects.filter(is_active=True).values_list("name", flat=True)
        )
        self.assertEqual(active_names, ["Empresa Ativa", "Empresa Legada"])
        self.assertFalse(ActiveCompany.objects.get(name="Empresa Inativa").is_active)
        self.assertFalse(ActiveCompany.objects.filter(name="Empresa fora do cadastro").exists())
        meeting = Meeting.objects.get()
        self.assertEqual(list(meeting.companies.values_list("name", flat=True)), ["Empresa Ativa"])

    def test_import_merges_metrics_without_deleting_or_duplicating(self):
        def make_upload(result):
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(["data", "indicador", "meta", "resultado", "status", "ordem"])
            worksheet.append(["2026-10-31", "Empresas ativas", "Manter", result, "Meta atendida", 1])
            payload = BytesIO()
            workbook.save(payload)
            return SimpleUploadedFile("dashboard.xlsx", payload.getvalue())

        self.assertEqual(import_dashboard_excel(make_upload("14")), 1)
        self.assertEqual(import_dashboard_excel(make_upload("14")), 0)
        self.assertEqual(
            list(DashboardMetric.objects.values_list("result_value", flat=True)),
            ["14"],
        )

        self.assertEqual(import_dashboard_excel(make_upload("17")), 1)
        self.assertEqual(
            list(DashboardMetric.objects.values_list("result_value", flat=True)),
            ["14", "17"],
        )

    def test_import_updates_existing_meeting_date_without_creating_duplicate(self):
        workbook = Workbook()
        report = workbook.active
        report.title = "RELATÓRIO TRIMESTRAL"
        report.append(["Indicador", "Resultado atual", "Meta / referência", "Situação"])
        report.append(["Empresas ativas", 1, "Manter", "Validar cadastro"])

        meetings = workbook.create_sheet("REUNIÕES")
        meetings.append(["ID reunião", "Data", "Tema/observações"])
        meetings.append(["R014", datetime(2026, 11, 12), "Reunião de teste"])

        def upload_workbook():
            payload = BytesIO()
            workbook.save(payload)
            return SimpleUploadedFile("nucleo.xlsx", payload.getvalue())

        import_dashboard_excel(upload_workbook())
        self.assertEqual(Meeting.objects.count(), 1)

        meetings["B2"] = datetime(2026, 11, 17)
        import_dashboard_excel(upload_workbook())

        meeting = Meeting.objects.get(title="R014 - Reunião de teste")
        self.assertEqual(Meeting.objects.count(), 1)
        self.assertEqual(meeting.date.isoformat(), "2026-11-17")

    def test_imports_flat_dashboard_template(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["data", "indicador", "meta", "resultado", "status", "ordem"])
        worksheet.append(["2026-10-31", "Indicador", "10", "12", "Meta superada", 1])

        payload = BytesIO()
        workbook.save(payload)
        upload = SimpleUploadedFile("dashboard.xlsx", payload.getvalue())

        self.assertEqual(import_dashboard_excel(upload), 1)
        metric = DashboardMetric.objects.get()
        self.assertEqual(metric.title, "Indicador")
        self.assertEqual(metric.status, DashboardMetric.STATUS_EXCEEDED)

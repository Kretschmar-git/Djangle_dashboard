import re
import unicodedata
from datetime import date, datetime
from io import BytesIO

from django.core.exceptions import ValidationError
from django.db import transaction
from openpyxl import load_workbook

from .models import ActiveCompany, DashboardMetric, Meeting, MeetingParticipant


def _normalize_text(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    return str(value).strip()


def _normalize_header(value):
    text = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode("ascii").lower()
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


HEADER_ALIASES = {
    "date": ("data", "date", "data reuniao", "data da reuniao"),
    "indicator": ("indicador", "indicator", "titulo", "nome", "indicador medida"),
    "target": ("meta", "meta referencia", "meta referencia inicial", "referencia inicial", "target"),
    "result": ("resultado", "resultado obtido", "resultado atual", "valor", "valor resultado", "actual", "result"),
    "status": ("status", "situacao", "situacao da meta", "estado", "status presenca"),
    "order": ("ordem", "order", "posicao", "sort order"),
    "company": ("empresa", "company"),
    "participant": ("nome", "participante", "participant"),
    "role": ("cargo", "funcao", "role"),
    "meeting_id": ("id reuniao", "id da reuniao", "meeting id"),
    "meeting_title": ("tema observacoes", "tema", "observacoes", "assunto"),
}


def _find_header(rows, required_fields):
    for row_index, row in enumerate(rows):
        indexes = {}
        for column_index, value in enumerate(row):
            header = _normalize_header(value)
            for field, aliases in HEADER_ALIASES.items():
                if header in aliases and field not in indexes:
                    indexes[field] = column_index
        if required_fields.issubset(indexes):
            return row_index, indexes
    return None, {}


def _parse_date(value):
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        try:
            from openpyxl.utils.datetime import from_excel

            parsed = from_excel(value)
        except (OverflowError, TypeError, ValueError):
            return None
        if isinstance(parsed, datetime):
            return parsed.date()
        return parsed if isinstance(parsed, date) else None

    text = str(value).strip()
    if not text:
        return None

    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%Y/%m/%d"]:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    # tenta lidar com datas em formato de texto do Excel
    cleaned = text.replace("/", "-")
    for fmt in ["%d-%m-%Y", "%Y-%m-%d", "%d-%m-%Y"]:
        try:
            return datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue

    return None


def _guess_status(value):
    if value is None:
        return DashboardMetric.STATUS_NA

    text = str(value).strip().lower()
    if not text or text in {"na", "n/a", "nan", "não aplicável", "nao aplicavel", "nao aplicavel"}:
        return DashboardMetric.STATUS_NA
    if "sem meta" in text or "não aplicável" in text or "nao aplicavel" in text:
        return DashboardMetric.STATUS_WARNING
    if text.startswith("meta superada") or text.startswith("superada"):
        return DashboardMetric.STATUS_EXCEEDED
    if text.startswith("sem meta"):
        return DashboardMetric.STATUS_WARNING
    if text.endswith("%"):
        try:
            pct = float(text.replace("%", "").replace(",", "."))
            return DashboardMetric.STATUS_EXCEEDED if pct > 100 else DashboardMetric.STATUS_ACHIEVED
        except ValueError:
            pass
    return DashboardMetric.STATUS_ACHIEVED


def normalize_status(value):
    text = _normalize_header(value)
    if not text:
        return DashboardMetric.STATUS_ACHIEVED
    if "nao atendid" in text or "abaixo da meta" in text:
        return DashboardMetric.STATUS_NOT_ACHIEVED
    if "superada" in text or "exced" in text:
        return DashboardMetric.STATUS_EXCEEDED
    if "nao aplicavel" in text or text == "na":
        return DashboardMetric.STATUS_NA
    if "sem meta" in text:
        return DashboardMetric.STATUS_WARNING
    if "atendida" in text or "meta" in text:
        return DashboardMetric.STATUS_ACHIEVED
    return _guess_status(text)


def _numeric_value(value):
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value or "").strip().replace("%", "").replace(" ", "")
    if not text:
        return None
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _format_metric_value(value, is_percentage=False):
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)):
        number = float(value)
        if is_percentage and abs(number) <= 1:
            number *= 100
        if number.is_integer():
            return f"{int(number)}" + ("%" if is_percentage else "")
        formatted = f"{number:.2f}".rstrip("0").rstrip(".").replace(".", ",")
        return formatted + ("%" if is_percentage else "")
    return _normalize_text(value)


def _report_metric_status(result, target, status):
    normalized_status = _normalize_header(status)
    if any(token in normalized_status for token in ("superada", "exced")):
        return DashboardMetric.STATUS_EXCEEDED
    if any(token in normalized_status for token in ("nao atendid", "abaixo da meta")):
        return DashboardMetric.STATUS_NOT_ACHIEVED
    if "nao aplicavel" in normalized_status or normalized_status == "na":
        return DashboardMetric.STATUS_NA
    if "sem meta" in normalized_status:
        return DashboardMetric.STATUS_WARNING
    if "atendida" in normalized_status:
        return DashboardMetric.STATUS_ACHIEVED

    result_number = _numeric_value(result)
    target_number = _numeric_value(target)
    if result_number is None:
        return DashboardMetric.STATUS_NA
    if target_number is not None:
        if result_number < target_number:
            return DashboardMetric.STATUS_NOT_ACHIEVED
        if result_number > target_number:
            return DashboardMetric.STATUS_EXCEEDED
        return DashboardMetric.STATUS_ACHIEVED
    return normalize_status(status or result)


def _report_reference_date(workbook):
    settings = workbook.get("INDICADORES")
    if settings is not None:
        rows = list(settings.iter_rows(values_only=True))
        year = None
        month = None
        for row in rows:
            if len(row) < 2:
                continue
            label = _normalize_header(row[0])
            if label == "data da ultima atualizacao":
                updated_at = _parse_date(row[1])
                if updated_at:
                    return updated_at
            if label == "ano de acompanhamento":
                year = _numeric_value(row[1])
            if label == "mes de referencia":
                month = _numeric_value(row[1])
        if year and month and 1 <= int(month) <= 12:
            return date(int(year), int(month), 1)
    return datetime.now().date()


def _attendance_rate(workbook):
    attendance_sheet = workbook.get("PRESENÇAS")
    if attendance_sheet is None:
        return None

    rows = list(attendance_sheet.iter_rows(values_only=True))
    header_row, indexes = _find_header(rows, {"status"})
    if header_row is None:
        return None

    status_index = indexes["status"]
    present = 0
    denominator = 0
    for row in rows[header_row + 1:]:
        status = _normalize_header(row[status_index] if status_index < len(row) else "")
        if status in {"presente", "ausente", "justificada"}:
            denominator += 1
            if status == "presente":
                present += 1

    return present / denominator if denominator else 0


def _sync_imported_entities(workbook):
    company_sheet = workbook.get("EMPRESAS")
    company_registry_imported = False
    if company_sheet is not None:
        rows = list(company_sheet.iter_rows(values_only=True))
        header_row, indexes = _find_header(rows, {"company"})
        if header_row is not None:
            company_registry_imported = True
            for row in rows[header_row + 1:]:
                company_index = indexes["company"]
                if company_index >= len(row) or row[company_index] is None:
                    continue
                company_name = str(row[company_index]).strip()
                if not company_name:
                    continue
                status_index = indexes.get("status")
                status_text = (
                    _normalize_header(row[status_index])
                    if status_index is not None and status_index < len(row)
                    else ""
                )
                is_active = not status_text or status_text in {"ativa", "ativo", "active", "sim", "yes"}
                ActiveCompany.objects.update_or_create(name=company_name, defaults={"is_active": is_active})

    participant_sheet = workbook.get("PARTICIPANTES")
    if participant_sheet is not None:
        rows = list(participant_sheet.iter_rows(values_only=True))
        header_row, indexes = _find_header(rows, {"participant"})
        if header_row is not None:
            for row in rows[header_row + 1:]:
                participant_index = indexes["participant"]
                if participant_index >= len(row) or row[participant_index] is None:
                    continue
                name = str(row[participant_index]).strip()
                if not name:
                    continue
                role_index = indexes.get("role")
                role = (
                    str(row[role_index]).strip()
                    if role_index is not None and role_index < len(row) and row[role_index] is not None
                    else ""
                )
                MeetingParticipant.objects.update_or_create(name=name, defaults={"role": role})

    meeting_ids = {}
    meeting_sheet = workbook.get("REUNIÕES")
    if meeting_sheet is not None:
        rows = list(meeting_sheet.iter_rows(values_only=True))
        header_row, indexes = _find_header(rows, {"meeting_id"})
        if header_row is not None:
            for row in rows[header_row + 1:]:
                meeting_index = indexes["meeting_id"]
                if meeting_index >= len(row) or row[meeting_index] is None:
                    continue
                meeting_id = str(row[meeting_index]).strip()
                if not meeting_id:
                    continue
                raw_title = row[indexes["meeting_title"]] if indexes.get("meeting_title", len(row)) < len(row) else None
                title = f"{meeting_id} - {str(raw_title).strip()}" if raw_title else f"Reunião {meeting_id}"
                date_index = indexes.get("date")
                meeting_date = (
                    _parse_date(row[date_index])
                    if date_index is not None and date_index < len(row)
                    else None
                ) or datetime.now().date()
                meeting, _ = Meeting.objects.update_or_create(
                    title=title,
                    defaults={"date": meeting_date},
                )
                meeting_ids[meeting_id] = meeting

    presence_sheet = workbook.get("PRESENÇAS")
    if presence_sheet is not None:
        rows = list(presence_sheet.iter_rows(values_only=True))
        header_row, indexes = _find_header(rows, {"meeting_id", "participant"})
        if header_row is not None:
            for row in rows[header_row + 1:]:
                meeting_index = indexes["meeting_id"]
                participant_index = indexes["participant"]
                meeting_id = str(row[meeting_index]).strip() if meeting_index < len(row) and row[meeting_index] else ""
                participant_name = (
                    str(row[participant_index]).strip()
                    if participant_index < len(row) and row[participant_index]
                    else ""
                )
                if not meeting_id or not participant_name:
                    continue

                company_index = indexes.get("company")
                company_name = (
                    str(row[company_index]).strip()
                    if company_index is not None and company_index < len(row) and row[company_index]
                    else ""
                )
                status_index = indexes.get("status")
                presence_status = (
                    _normalize_header(row[status_index])
                    if status_index is not None and status_index < len(row)
                    else ""
                )
                if presence_status not in {"presente", "x"}:
                    continue

                meeting = meeting_ids.get(meeting_id)
                if not meeting_id or not participant_name:
                    continue
                if meeting is None:
                    date_index = indexes.get("date")
                    meeting_date = (
                        _parse_date(row[date_index])
                        if date_index is not None and date_index < len(row)
                        else None
                    ) or datetime.now().date()
                    meeting, _ = Meeting.objects.get_or_create(
                        title=f"Reunião {meeting_id}",
                        date=meeting_date,
                    )
                    meeting_ids[meeting_id] = meeting
                participant, _ = MeetingParticipant.objects.get_or_create(name=participant_name)
                meeting.participants.add(participant)
                if company_name:
                    if company_registry_imported:
                        company = ActiveCompany.objects.filter(name=company_name, is_active=True).first()
                        if company is not None:
                            meeting.companies.add(company)
                    else:
                        company, _ = ActiveCompany.objects.get_or_create(
                            name=company_name,
                            defaults={"is_active": True},
                        )
                        meeting.companies.add(company)


def import_dashboard_excel(file):
    if file is None:
        raise ValidationError("Arquivo não informado.")

    dataset = file.read()
    if not dataset:
        raise ValidationError("Arquivo vazio.")

    workbook = load_workbook(BytesIO(dataset), data_only=True)
    workbook_by_name = {sheet.title: sheet for sheet in workbook.worksheets}
    report_sheet = workbook_by_name.get("RELATÓRIO TRIMESTRAL")
    candidate_sheets = ([report_sheet] if report_sheet is not None else []) + [
        sheet for sheet in workbook.worksheets if sheet is not report_sheet
    ]

    worksheet = None
    rows = []
    header_row = None
    indexes = {}
    for candidate in candidate_sheets:
        candidate_rows = list(candidate.iter_rows(values_only=True))
        candidate_header, candidate_indexes = _find_header(candidate_rows, {"indicator", "result"})
        if candidate_header is not None:
            worksheet = candidate
            rows = candidate_rows
            header_row = candidate_header
            indexes = candidate_indexes
            break

    if worksheet is None or header_row is None:
        raise ValidationError(
            "Não encontrei a tabela de indicadores. Use as colunas Indicador e Resultado "
            "ou a aba RELATÓRIO TRIMESTRAL do modelo completo."
        )

    is_report = worksheet.title == "RELATÓRIO TRIMESTRAL"
    reference_date = _report_reference_date(workbook_by_name) if is_report else None
    attendance_rate = _attendance_rate(workbook_by_name) if is_report else None
    imported_metrics = []
    for row_index, row in enumerate(rows[header_row + 1:], start=1):
        if not row or not any(cell is not None and str(cell).strip() for cell in row):
            continue

        def value_for(key):
            idx = indexes.get(key)
            return row[idx] if idx is not None and idx < len(row) else None

        raw_date = value_for("date")
        raw_title = value_for("indicator")
        raw_target = value_for("target")
        raw_result = value_for("result")
        raw_status = value_for("status")
        sort_order = value_for("order")
        title = _normalize_text(raw_title)

        if not title:
            continue

        if is_report and raw_target is None and raw_status is None:
            continue

        normalized_title = _normalize_header(title)
        if "presenca" in normalized_title and attendance_rate is not None:
            raw_result = attendance_rate
        is_percentage = is_report and any(
            token in normalized_title for token in ("presenca", "execucao", "percentual", "taxa")
        )
        target_text = _format_metric_value(raw_target, is_percentage) or "Referência não informada"
        result_value = _format_metric_value(raw_result, is_percentage) or "NA"
        parsed_date = _parse_date(raw_date) or reference_date or datetime.now().date()

        try:
            order_number = int(str(sort_order).strip()) if sort_order is not None else row_index
        except ValueError:
            order_number = row_index

        status = (
            _report_metric_status(raw_result, raw_target, raw_status)
            if is_report
            else normalize_status(_normalize_text(raw_status) or result_value)
        )
        imported_metrics.append(
            DashboardMetric(
                date=parsed_date,
                title=title,
                target_text=target_text,
                result_value=result_value,
                status=status,
                sort_order=order_number,
            )
        )

    if not imported_metrics:
        raise ValidationError("A tabela de indicadores não contém dados para importar.")

    new_metrics = []
    with transaction.atomic():
        _sync_imported_entities(workbook_by_name)
        existing_metric_keys = set(
            DashboardMetric.objects.values_list(
                "date",
                "title",
                "target_text",
                "result_value",
                "status",
            )
        )
        for metric in imported_metrics:
            metric_key = (
                metric.date,
                metric.title,
                metric.target_text,
                metric.result_value,
                metric.status,
            )
            if metric_key in existing_metric_keys:
                continue
            new_metrics.append(metric)
            existing_metric_keys.add(metric_key)

        if new_metrics:
            DashboardMetric.objects.bulk_create(new_metrics)

    return len(new_metrics)

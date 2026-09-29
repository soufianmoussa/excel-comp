"""Create safe downloadable workbooks in memory."""
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from comparator import STATUSES


SHEETS = {"MATCH": "Correspondances", "DIFFERENT": "Différences", "MISSING_IN_B": "Absents",
          "INCOMPLETE": "Incomplets", "DUPLICATE": "Doublons"}


def write_rows(sheet, rows):
    for row in rows:
        sheet.append(list(row))
        for cell in sheet[sheet.max_row]:
            # Excel must never execute uploaded text as a formula.
            if isinstance(cell.value, str):
                cell.data_type = "s"
                cell.number_format = "@"
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for cell in sheet[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = PatternFill("solid", fgColor="176B67")
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for i, column in enumerate(sheet.columns, 1):
        width = max((len(str(c.value or "")) for c in column), default=12)
        sheet.column_dimensions[get_column_letter(i)].width = min(55, max(16, width + 2))


def export_excel(results, metadata=None):
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Résumé"
    counts = results["Code statut"].value_counts() if not results.empty else {}
    rows = [["Indicateur", "Valeur"], ["Total analysé", len(results)]]
    rows.extend([STATUSES[key], int(counts.get(key, 0))] for key in STATUSES)
    rows.extend([key, str(value)] for key, value in (metadata or {}).items())
    write_rows(summary, rows)
    for status, title in SHEETS.items():
        sheet = workbook.create_sheet(title)
        if results.empty:
            write_rows(sheet, [["Statut"]])
        else:
            subset = results.loc[results["Code statut"] == status].drop(columns="Code statut")
            write_rows(sheet, [subset.columns.tolist(), *subset.itertuples(index=False, name=None)])
    output = BytesIO()
    workbook.save(output)
    return output.getvalue()

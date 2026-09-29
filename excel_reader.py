"""Read administrative workbooks in memory, retaining source-row provenance."""
from dataclasses import dataclass, field
from io import BytesIO
import re
import warnings
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

from normalizer import build_rib, name, text, rib


class WorkbookError(ValueError):
    pass


@dataclass
class Record:
    row: int
    values: dict
    issues: dict = field(default_factory=dict)


@dataclass
class Table:
    sheet: str
    header: int
    end_header: int
    columns: dict
    records: list = field(default_factory=list)

    @property
    def available(self):
        cols = self.columns
        return {f for f, ok in {
            "name": "name" in cols or {"first", "last"} <= cols.keys(),
            "ddp": "ddp" in cols, "cin": "cin" in cols,
            "rib": "rib" in cols or {"bank", "city", "account", "key"} <= cols.keys(),
        }.items() if ok}


def header_key(value):
    # Some source documents literally contain the four characters \\xb0.
    return re.sub(r"[^A-Z0-9]", "", name(value).replace("\\XB0", ""))


ALIASES = {
    "name": ["NOM ET PRENOM", "NOM PRENOM", "NOM ET PRENOMS"],
    "last": ["NOM"], "first": ["PRENOM", "PRENOMS"],
    "ddp": ["N DDP", "DDP", "DDR", "N DDR", "NUMERO DDP"],
    "cin": ["CIN", "N CIN", "NUMERO CIN"], "rib": ["RIB", "R.I.B"],
    "bank": ["BANQUE", "CODE BANQUE"], "city": ["VILLE", "CODE VILLE"],
    "account": ["NUMERO COMPTE RIB", "NUMORO COMPTE RIB", "NUMEORO COMPTE RIB", "NUMEORO COMPTE RIBE", "COMPTE RIB", "NUMERO DE COMPTE RIB"],
    "key": ["CLE", "CLE RIB"],
}
LOOKUP = {header_key(alias): field for field, aliases in ALIASES.items() for alias in aliases}


def open_workbook(data):
    try:
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", message="Cannot parse header or footer.*")
            return load_workbook(BytesIO(data), data_only=False)
    except (BadZipFile, InvalidFileException, OSError, ValueError, KeyError, TypeError) as exc:
        raise WorkbookError("Impossible de lire ce fichier. Choisissez un classeur .xlsx valide, non protégé par mot de passe.") from exc


def cell_at(sheet, row, column):
    cell = sheet.cell(row, column)
    if cell.value is None:
        for merged in sheet.merged_cells.ranges:
            if merged.min_row <= row <= merged.max_row and merged.min_col <= column <= merged.max_col:
                return sheet.cell(merged.min_row, merged.min_col)
    return cell


def detect_header(sheet, row):
    columns = {}
    end = row
    for col in range(1, sheet.max_column + 1):
        cell = cell_at(sheet, row, col)
        key = LOOKUP.get(header_key(cell.value))
        if key and key not in columns:
            columns[key] = cell.column
            for merged in sheet.merged_cells.ranges:
                if cell.coordinate in merged:
                    end = max(end, merged.max_row)
    return Table(sheet.title, row, end, columns)


def detect_tables(workbook):
    tables = []
    for sheet in workbook:
        candidates = [detect_header(sheet, row) for row in range(1, min(sheet.max_row, 50) + 1)]
        candidates = [t for t in candidates if len(t.available) >= 2]
        if candidates:
            tables.append(max(candidates, key=lambda t: (len(t.available), len(t.columns), -t.header)))
    return tables


def read_cell(cell):
    value = cell.value
    if cell.data_type == "f":
        return text(value), "Formule : remplacez-la par sa valeur dans Excel."
    if cell.data_type == "e":
        return text(value), "Erreur Excel : corrigez la cellule source."
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if abs(value) >= 10 ** 15:
            return text(value), "Précision incertaine : identifiant numérique de plus de 15 chiffres. Vérifiez la source et enregistrez-le comme texte."
        if isinstance(value, float) and not value.is_integer():
            return text(value), "Identifiant numérique contenant des décimales."
        if re.fullmatch(r"0+", cell.number_format):
            return text(value).zfill(len(cell.number_format)), ""
    return text(value), ""


def read_table(workbook, table):
    sheet = workbook[table.sheet]
    records = []
    for row in range(table.end_header + 1, sheet.max_row + 1):
        raw, errors = {}, {}
        for key, col in table.columns.items():
            raw[key], error = read_cell(cell_at(sheet, row, col))
            if error:
                errors[key] = error
        if not any(v.strip() for v in raw.values()):
            continue
        # Repeated page headers are not people.
        if sum(LOOKUP.get(header_key(v)) == k for k, v in raw.items() if v) >= 2:
            continue
        values = {
            "name": raw.get("name", " ".join(raw.get(k, "") for k in ("first", "last"))),
            "ddp": raw.get("ddp", ""), "cin": raw.get("cin", ""),
            "rib": raw.get("rib", build_rib(*(raw.get(k, "") for k in ("bank", "city", "account", "key")))),
        }
        issues = {}
        dependencies = {"name": ["name"] if "name" in raw else ["first", "last"],
                        "ddp": ["ddp"], "cin": ["cin"],
                        "rib": ["rib"] if "rib" in raw else ["bank", "city", "account", "key"]}
        for key, sources in dependencies.items():
            messages = [errors[s] for s in sources if s in errors]
            if any(not raw.get(s, "").strip() for s in sources):
                messages.append("Information manquante.")
            if messages:
                issues[key] = " ".join(dict.fromkeys(messages))
        if "rib" not in raw:
            for key, width in zip(("bank", "city", "account", "key"), (3, 3, 16, 2)):
                part = rib(raw.get(key, ""))
                if part and (not part.isascii() or not part.isdigit() or len(part) > width):
                    issues["rib"] = "Composant RIB invalide (longueurs attendues : 3 / 3 / 16 / 2)."
        full = rib(values["rib"])
        if full and (not full.isascii() or not full.isdigit() or len(full) != 24):
            issues["rib"] = "RIB invalide : 24 chiffres attendus. " + issues.get("rib", "")
        records.append(Record(row, values, issues))
    return Table(table.sheet, table.header, table.end_header, table.columns, records)

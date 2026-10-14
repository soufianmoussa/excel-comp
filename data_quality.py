"""Independent, coordinate-preserving quality checks. No guessed business errors."""
from collections import defaultdict
from dataclasses import asdict, dataclass, replace
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import re
import statistics

import pandas as pd

from excel_reader import LOOKUP, header_key
from normalizer import cin, identifier, name, rib, text
from row_classifier import SUMMARY_TYPES, parse_summary_amount


SEVERITIES = ["Erreur", "Avertissement", "Information"]
FINDING_COLUMNS = ["Fichier", "Feuille", "Ligne", "Colonne", "Titre", "Cellule", "Valeur originale", "Type", "Gravité", "Explication", "Règle appliquée", "Contexte de validation"]
SEMANTICS = {
    "start": ["DEBUT", "DATE DEBUT", "DATE DE DEBUT"],
    "end": ["FIN", "DATE FIN", "DATE DE FIN"],
    "amount_net": ["NET(DH)", "MONTANT NET", "NET A PAYER"],
    "amount": ["BRUT", "RETENUES", "MONTANT", "TAUX"],
    "category": ["CATEGORIE", "GRADE/ECHELLE"],
}
SEMANTIC_LOOKUP = {header_key(v): key for key, values in SEMANTICS.items() for v in values}


def semantic(title):
    return LOOKUP.get(header_key(title), SEMANTIC_LOOKUP.get(header_key(title), ""))


@dataclass
class Rule:
    required: bool = False
    kind: str = "auto"
    min_length: int | None = None
    max_length: int | None = None
    allowed: str = ""
    minimum: str = ""
    maximum: str = ""
    date_min: str = ""
    date_max: str = ""
    custom: str = ""


def parse_number(value):
    if isinstance(value, bool) or isinstance(value, (date, datetime)):
        raise ValueError("not a number")
    raw = text(value).strip().replace("\u00a0", "").replace("\u202f", "").replace(" ", "")
    if not re.fullmatch(r"[+-]?(?:\d+(?:[.,]\d+)?|[.,]\d+)", raw):
        raise ValueError("not an unambiguous number")
    try:
        result = Decimal(raw.replace(",", "."))
    except InvalidOperation as exc:
        raise ValueError("not a number") from exc
    if not result.is_finite():
        raise ValueError("not finite")
    return result


def parse_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text(value).strip(), fmt).date()
        except ValueError:
            pass
    raise ValueError("not a date")


def default_rule(title):
    kind = semantic(title)
    rule = Rule()
    if kind in {"name", "first", "last"}:
        rule.kind, rule.required = "text", True
    elif kind in {"cin", "ddp", "rib", "bank", "city", "account", "key"}:
        rule.kind, rule.required = "identifier", True
    elif kind in {"start", "end"}:
        rule.kind = "date"
    elif kind in {"amount", "amount_net"}:
        rule.kind = "number"
    elif kind == "category":
        rule.kind = "text"
    return rule


def infer_kind(table, col):
    data_rows = set(table.rows)
    values = [c.value for (row, column), c in table.cells.items()
              if column == col and row in data_rows and c.value is not None and c.data_type != "f"][:300]
    if not values:
        return "text"
    if any(isinstance(v, str) and re.fullmatch(r"0[0-9]+", v.strip()) for v in values):
        return "identifier"
    for kind, parser in (("date", parse_date), ("number", parse_number)):
        valid = 0
        for value in values:
            try:
                parser(value)
                valid += 1
            except ValueError:
                pass
        if valid / len(values) >= .9:
            return kind
    return "text"


def validate_rule(rule):
    if rule.kind not in {"auto", "text", "number", "date", "identifier"}:
        raise ValueError("Type de règle non reconnu.")
    for value in (rule.min_length, rule.max_length):
        if value is not None and (value < 0 or int(value) != value):
            raise ValueError("Les longueurs doivent être des entiers positifs ou nuls.")
    if rule.min_length is not None and rule.max_length is not None and rule.min_length > rule.max_length:
        raise ValueError("La longueur minimale dépasse la longueur maximale.")
    try:
        low = parse_number(rule.minimum) if rule.minimum else None
        high = parse_number(rule.maximum) if rule.maximum else None
        dlow = parse_date(rule.date_min) if rule.date_min else None
        dhigh = parse_date(rule.date_max) if rule.date_max else None
    except ValueError as exc:
        raise ValueError("Une borne numérique ou une date de la configuration est invalide.") from exc
    if low is not None and high is not None and low > high:
        raise ValueError("Le minimum numérique dépasse le maximum.")
    if dlow and dhigh and dlow > dhigh:
        raise ValueError("La date minimale dépasse la date maximale.")
    if rule.custom and (":" not in rule.custom or rule.custom.split(":", 1)[0] not in {"prefixe", "suffixe", "contient", "sans"} or not rule.custom.split(":", 1)[1]):
        raise ValueError("Règle personnalisée : utilisez prefixe:texte, suffixe:texte, contient:texte ou sans:texte.")


def finding(label, table, row, col, category, severity, explanation):
    cell = table.cells[row, col]
    classification = table.row_classifications.get(row)
    context = f"{classification.kind} ; rôle {cell.role} : {classification.reason}" if classification else "DATA : cellule d'un enregistrement individuel."
    return dict(zip(FINDING_COLUMNS, [label, table.sheet, row, col, table.row_schemas.get(row, table).all_columns[col],
        cell.coordinate, cell.display, category, severity, explanation, category, context]))


def analyze_table(table, label, rules=None, summary_rules=None):
    """Validate each data region under its own headers, summaries separately."""
    rules = rules or {}
    groups = defaultdict(list)
    for row in table.rows:
        schema = table.row_schemas.get(row, table)
        groups[schema.header].append(row)
    frames = []
    for rows in groups.values():
        schema = table.row_schemas.get(rows[0], table)
        subset = replace(table, rows=rows, all_columns=schema.all_columns, columns=schema.columns)
        # Admin overrides target physical columns; defaults follow each region's header.
        local_rules = {col: rule for col, rule in rules.items()
                       if col in schema.all_columns and (schema.header == table.header or
                       header_key(schema.all_columns[col]) == header_key(table.all_columns[col]) or
                       rule != default_rule(table.all_columns[col]))}
        frames.append(_analyze_data(subset, label, local_rules))
    summary_findings = []
    for row, classification in table.row_classifications.items():
        if classification.kind not in SUMMARY_TYPES:
            continue
        for col in classification.amount_columns:
            cell = table.cells[row, col]
            if cell.role == "MERGED_PLACEHOLDER" or cell.data_type == "f":
                continue  # No recalculation is claimed without a calculation rule.
            try:
                amount = parse_summary_amount(cell.value)
                rule = (summary_rules or {}).get(col)
                if rule:
                    validate_rule(rule)
                    if (rule.minimum and amount < parse_number(rule.minimum)) or (rule.maximum and amount > parse_number(rule.maximum)):
                        summary_findings.append(finding(label, table, row, col, "Plage du montant de synthèse", "Erreur", "Le montant de synthèse dépasse ses bornes propres configurées."))
            except ValueError:
                summary_findings.append(finding(label, table, row, col, "Montant de synthèse invalide", "Erreur", "Cette cellule a le rôle de montant de synthèse ; son écriture n'est pas reconnue."))
    frames.append(pd.DataFrame(summary_findings, columns=FINDING_COLUMNS))
    return pd.concat(frames, ignore_index=True).drop_duplicates() if frames else pd.DataFrame(columns=FINDING_COLUMNS)


def _analyze_data(table, label, rules=None):
    rules = rules or {}
    findings = []
    def add(row, col, category, severity, message):
        findings.append(finding(label, table, row, col, category, severity, message))
    for col, title in table.all_columns.items():
        rule = rules.get(col, default_rule(title))
        validate_rule(rule)
        sem = semantic(title)
        kind = infer_kind(table, col) if rule.kind == "auto" else rule.kind
        strict = "Avertissement" if rule.kind == "auto" else "Erreur"
        variants, numbers, formats = defaultdict(list), [], defaultdict(list)
        casing = defaultdict(list)
        allowed = {v.strip() for v in rule.allowed.split("|") if v.strip()}
        for row in table.rows:
            cell = table.cells[row, col]
            if cell.role == "MERGED_PLACEHOLDER":
                continue
            value, raw = cell.value, cell.display
            if value is None or not raw.strip():
                if rule.required:
                    add(row, col, "Valeur manquante", "Erreur", "Information obligatoire manquante.")
                continue
            if cell.data_type == "f":
                add(row, col, "Formule non évaluée", "Avertissement", "Cette formule n'est pas calculée par l'application. Fournissez sa valeur pour vérifier et comparer cette cellule.")
                continue
            if cell.data_type == "e":
                add(row, col, "Erreur Excel", "Erreur", "La cellule contient une erreur Excel.")
                continue
            if isinstance(value, str):
                if value != value.strip():
                    add(row, col, "Espaces", "Information", "Espaces en début ou en fin de valeur ; valeur originale conservée.")
                if any(ord(c) < 32 and c not in "\n\r\t" for c in value):
                    add(row, col, "Caractères inattendus", "Avertissement", "Caractère de contrôle inhabituel à vérifier.")
                variants[name(value)].append((row, value))
                stripped = value.strip()
                if kind == "text" and any(c.isalpha() for c in stripped):
                    style = "majuscules" if stripped.isupper() else "minuscules" if stripped.islower() else "initiales" if stripped.istitle() else "mixte"
                    casing[style].append(row)
            normalized = cin(raw) if sem == "cin" else identifier(raw) if kind == "identifier" else raw.strip()
            length_value = normalized
            if rule.min_length is not None and len(length_value) < rule.min_length or rule.max_length is not None and len(length_value) > rule.max_length:
                add(row, col, "Longueur", "Erreur", "Longueur hors des limites configurées.")
            if allowed and raw.strip() not in allowed:
                add(row, col, "Valeur autorisée", "Erreur", "Valeur absente de la liste autorisée configurée.")
            if rule.custom:
                operation, argument = rule.custom.split(":", 1)
                valid = {"prefixe": raw.startswith(argument), "suffixe": raw.endswith(argument),
                         "contient": argument in raw, "sans": argument not in raw}[operation]
                if not valid:
                    add(row, col, "Règle personnalisée", "Erreur", "La valeur ne respecte pas la règle personnalisée configurée.")
            if kind == "text" and rule.kind != "auto" and not isinstance(value, str):
                add(row, col, "Type", strict, "Un texte est attendu dans cette colonne.")
            if kind == "number" or rule.minimum or rule.maximum:
                try:
                    number = parse_number(value)
                    numbers.append((row, number))
                    if ((rule.minimum and number < parse_number(rule.minimum)) or
                            (rule.maximum and number > parse_number(rule.maximum))):
                        add(row, col, "Plage numérique", "Erreur", "Nombre hors de la plage configurée.")
                except ValueError:
                    add(row, col, "Type", strict, "Nombre attendu ; valeur non reconnue ou ambiguë.")
            if kind == "date" or rule.date_min or rule.date_max:
                try:
                    day = parse_date(value)
                    if (rule.date_min and day < parse_date(rule.date_min)) or (rule.date_max and day > parse_date(rule.date_max)):
                        add(row, col, "Plage de dates", "Erreur", "Date hors de la période configurée.")
                    fmt = cell.number_format if isinstance(value, (date, datetime)) else re.sub(r"\d", "d", raw.strip())
                    formats[fmt].append(row)
                except ValueError:
                    add(row, col, "Date invalide", strict, "Date non reconnue. Formats acceptés : jour/mois/année ou année-mois-jour, notamment.")
            if kind == "identifier":
                if isinstance(value, (int, float)) and not isinstance(value, bool) and abs(value) >= 10**15:
                    add(row, col, "Précision numérique", "Avertissement", "Identifiant numérique de plus de 15 chiffres : des chiffres ont pu être perdus dans Excel. Vérifiez la source.")
                elif isinstance(value, (date, bool)) or isinstance(value, float) and not value.is_integer():
                    add(row, col, "Type", "Erreur", "Identifiant attendu, sans décimales ni date.")
                if sem in {"bank", "city", "account", "key", "rib"}:
                    width = {"bank": 3, "city": 3, "account": 16, "key": 2, "rib": 24}[sem]
                    digits = rib(raw)
                    if not digits.isascii() or not digits.isdigit() or len(digits) > width or sem == "rib" and len(digits) != width:
                        add(row, col, "Format RIB", "Erreur", f"Chiffres uniquement ; longueur attendue : {width}. Les composants courts peuvent être complétés par des zéros.")
                    elif len(digits) < width:
                        add(row, col, "Format RIB", "Information", f"Composant complété à gauche par des zéros sur {width} chiffres pour la comparaison.")
                elif sem == "cin" and not re.fullmatch(r"[A-Z]+[0-9]+", cin(raw)):
                    add(row, col, "Format identifiant", "Avertissement", "Format CIN inhabituel (lettres suivies de chiffres habituellement) : vérifiez sans conclure à une erreur.")
                elif sem == "ddp" and not re.fullmatch(r"[\w/-]+", normalized):
                    add(row, col, "Format identifiant", "Avertissement", "Caractères inhabituels dans cet identifiant.")
            if sem in {"name", "first", "last"} and any(not (c.isalpha() or c.isspace() or c in "-'’.") for c in raw):
                add(row, col, "Caractères inattendus", "Avertissement", "Le nom contient des caractères inhabituels : vérifiez la source.")
        case_flagged = set()
        for items in variants.values():
            if len({value.strip() for _, value in items}) > 1:
                for row, _ in items:
                    add(row, col, "Casse / écriture", "Information", "Variantes de casse, d'accents ou d'espacement pour une même valeur normalisée.")
                    case_flagged.add(row)
        if len(casing) > 1:
            dominant = max(casing, key=lambda style: len(casing[style]))
            for style, rows in casing.items():
                if style != dominant:
                    for row in rows:
                        if row not in case_flagged:
                            add(row, col, "Casse / écriture", "Information", "Présentation en majuscules/minuscules différente d'autres valeurs de la colonne ; aucune faute n'est présumée.")
        if len(formats) > 1:
            for rows in formats.values():
                for row in rows:
                    add(row, col, "Format de date", "Information", "Plusieurs présentations de dates sont utilisées dans cette colonne.")
        if len(numbers) >= 8:
            sorted_values = sorted(v for _, v in numbers)
            q1, _, q3 = statistics.quantiles(sorted_values, n=4, method="inclusive")
            spread = q3 - q1
            if spread > 0:
                for row, value in numbers:
                    if value < q1 - 3 * spread or value > q3 + 3 * spread:
                        add(row, col, "Nombre atypique", "Avertissement", "Valeur éloignée des autres nombres de la colonne ; aucune erreur métier n'est présumée.")
    starts = [c for c, t in table.all_columns.items() if semantic(t) == "start"]
    ends = [c for c, t in table.all_columns.items() if semantic(t) == "end"]
    if len(starts) == len(ends) == 1:
        for row in table.rows:
            try:
                if parse_date(table.cells[row, ends[0]].value) < parse_date(table.cells[row, starts[0]].value):
                    for col in (starts[0], ends[0]):
                        add(row, col, "Chronologie", "Erreur", "La date de fin précède la date de début sur cette ligne.")
            except ValueError:
                pass
    return pd.DataFrame(findings, columns=FINDING_COLUMNS).drop_duplicates(subset=["Fichier", "Feuille", "Cellule", "Type", "Explication"])


def rules_frame(tables):
    rows = []
    for table in tables:
        for col, title in table.all_columns.items():
            rows.append({"Feuille": table.sheet, "Colonne": col, "Titre": title,
                         "Type suggéré": infer_kind(table, col), **asdict(default_rule(title))})
    return pd.DataFrame(rows)


def frame_rules(frame):
    result = {}
    for row in frame.to_dict("records"):
        values = {}
        for key in Rule.__dataclass_fields__:
            value = row[key]
            if key in {"min_length", "max_length"}:
                value = None if pd.isna(value) or value == "" else int(value)
            elif key == "required":
                value = bool(value)
            else:
                value = "" if pd.isna(value) else str(value)
            values[key] = value
        rule = Rule(**values)
        validate_rule(rule)
        result[row["Feuille"], int(row["Colonne"])] = rule
    return result

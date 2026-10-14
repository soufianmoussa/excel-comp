"""Conservative row classification using physical cells, layout and neighbours."""
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import re

from normalizer import name, text


SUMMARY_TYPES = {"REPORT", "TOTAL", "SUBTOTAL", "SUMMARY"}
AMOUNT_HEADERS = {"NETDH", "MONTANTNET", "NETAPAYER", "MONTANT", "BRUT", "RETENUES", "TAUX", "NOMBRE"}
SUMMARY_PREFIX = re.compile(r"^(A REPORTER|REPORT|TOTAL GENERAL|MONTANT TOTAL|SOUS[ -]TOTAL|TOTAL|SOMME|CUMUL)(?=$|[\s.:…-])")


@dataclass
class RowClassification:
    kind: str
    reason: str
    amount_columns: tuple = ()
    confidence: float = 0.95


def summary_label(value):
    normalized = name(value)
    match = SUMMARY_PREFIX.match(normalized)
    if not match:
        return None
    suffix = normalized[match.end():].lstrip(" .:…-\t")
    if suffix:
        try:
            parse_summary_amount(suffix)
        except ValueError:
            return None  # A person such as "Total Martin" is not a total.
    word = match.group(1)
    kind = "REPORT" if word in {"REPORT", "A REPORTER"} else "SUBTOTAL" if word.startswith("SOUS") else "TOTAL"
    return kind, suffix


def parse_summary_amount(value):
    """Parse printed summary amounts only; never change identifier parsing."""
    if isinstance(value, bool):
        raise ValueError("not an amount")
    raw = name(value).strip()
    prefix = SUMMARY_PREFIX.match(raw)
    if prefix:
        raw = raw[prefix.end():].lstrip(" .:…-\t")
    raw = re.sub(r"\s*(DHS?\.?|MAD|DIRHAMS?)\s*$", "", raw)
    raw = raw.replace("\u00a0", " ").replace("\u202f", " ").strip()
    # French grouping: 15.147,60 / 15 147,60. Inconsistent grouping is rejected.
    if re.fullmatch(r"[+-]?\d{1,3}(?:\.\d{3})+(?:,\d+)?", raw):
        raw = raw.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"[+-]?\d{1,3}(?: \d{3})+(?:[,.]\d+)?", raw):
        raw = raw.replace(" ", "").replace(",", ".")
    elif re.fullmatch(r"[+-]?\d+(?:[,.]\d+)?", raw):
        raw = raw.replace(",", ".")
    else:
        raise ValueError("not an unambiguous summary amount")
    try:
        number = Decimal(raw)
    except InvalidOperation as exc:
        raise ValueError("not an amount") from exc
    if not number.is_finite():
        raise ValueError("not finite")
    return number


def classify_rows(sheet, table, header_lookup, key_function, schemas=None):
    """A structural label alone never removes a row from record validation."""
    physical = {row: {col: sheet.cell(row, col) for col in table.all_columns
                      if sheet.cell(row, col).value is not None and text(sheet.cell(row, col).value).strip()}
                for row in range(1, sheet.max_row + 1)}
    merges = list(sheet.merged_cells.ranges)
    numeric_columns = {col for col, title in table.all_columns.items() if key_function(title) in AMOUNT_HEADERS}

    def has_identifier(row):
        schema = (schemas or {}).get(row, table)
        for role in ("cin", "ddp", "rib", "account"):
            col = schema.columns.get(role)
            cell = physical.get(row, {}).get(col)
            if cell is None or cell.data_type in {"f", "e"}:
                continue
            value = name(cell.value).replace(" ", "")
            if role == "cin" and re.fullmatch(r"[A-Z]+\d+", value):
                return True
            if role in {"ddp", "account", "rib"} and re.fullmatch(r"\d+(?:\.0+)?", value):
                return True
        return False

    identified = {row for row in physical if has_identifier(row)}
    neighbour_styles = {sheet.cell(r, c).style_id for r in identified for c in physical[r]}
    classifications = {}
    for row, occupied in physical.items():
        if not occupied:
            classifications[row] = RowClassification("EMPTY", "Aucune valeur physique sur cette ligne.")
            continue
        row_schema = (schemas or {}).get(row, table)
        if row_schema.header <= row <= row_schema.end_header:
            classifications[row] = RowClassification("HEADER", "Ligne appartenant aux en-têtes détectés.")
            continue
        header_hits = sum(header_lookup.get(key_function(cell.value)) == role
                          for role, col in table.columns.items() if (cell := occupied.get(col)) is not None)
        if header_hits >= 2:
            classifications[row] = RowClassification("HEADER", "Plusieurs intitulés reconnus aux positions des colonnes : en-tête répété.")
            continue
        # Identify people using the active header layout, including later tables.
        labels = {col: result for col, cell in occupied.items() if (result := summary_label(cell.value))}
        local_identity = any((cell := occupied.get(col)) is not None and not summary_label(cell.value)
                             and re.fullmatch(r"[A-Z]*\d+(?:\.0+)?", name(cell.value).replace(" ", ""))
                             and (not labels or role == "cin" and re.fullmatch(r"[A-Z]+\d+", name(cell.value).replace(" ", ""))
                                  or role == "rib" and re.fullmatch(r"\d{24}", name(cell.value).replace(" ", "")))
                             for role, col in row_schema.columns.items() if role in {"cin", "ddp", "rib", "account"})
        if local_identity:
            classifications[row] = RowClassification("DATA", "Identifiant individuel présent ; contrôles de données conservés.")
            continue
        spanning = [m for m in merges if m.min_row <= row <= m.max_row and m.max_col > m.min_col]
        # A two-column merged name in a wide record is not a presentation line.
        broad_merge = any(m.max_col - m.min_col + 1 >= max(3, len(table.all_columns) // 2) for m in spanning)
        sparse = len(occupied) <= max(2, len(table.all_columns) // 3)
        neighbour_data = any(n in identified for n in range(max(table.end_header + 1, row - 2), min(sheet.max_row, row + 2) + 1) if n != row)
        styled = any(cell.style_id not in neighbour_styles and (cell.font.bold or cell.font.italic or
                     cell.alignment.horizontal in {"center", "right"} or cell.border.top.style or cell.border.bottom.style)
                     for cell in occupied.values()) if neighbour_styles else False
        numeric_columns = {col for col, title in row_schema.all_columns.items() if key_function(title) in AMOUNT_HEADERS}
        amounts = set()
        for col, cell in occupied.items():
            if col in labels and not labels[col][1]:
                continue
            try:
                parse_summary_amount(cell.value)
                amounts.add(col)
            except ValueError:
                if cell.data_type == "f" or col in numeric_columns and col not in labels:
                    amounts.add(col)
        if labels and (amounts or broad_merge and sparse and neighbour_data):
            first = min(labels)
            classifications[row] = RowClassification(labels[first][0],
                "Libellé de synthèse, absence d'identifiant et montant/position ou mise en page de synthèse concordants.", tuple(sorted(amounts)))
            continue
        if row < table.header:
            classifications[row] = RowClassification("TITLE", "Texte situé avant les en-têtes du tableau.")
            continue
        note_cue = any(re.match(r"^(NOTE|OBSERVATION|REMARQUE|COMMENTAIRE|NB)\b", name(c.value)) for c in occupied.values())
        if note_cue and sparse and (broad_merge or styled) and neighbour_data:
            classifications[row] = RowClassification("NOTES", "Libellé de note, présentation distincte et absence d'identifiant.", confidence=.85)
            continue
        if not labels and amounts and broad_merge and sparse and styled and neighbour_data:
            classifications[row] = RowClassification("SUMMARY", "Montant global sans identifiant, fusion large, style distinct et voisinage de données.", tuple(sorted(amounts)), .8)
            continue
        if broad_merge and sparse:
            title_cue = any(re.match(r"^(GROUPE|SECTION|LISTE|ETAT|BENEFICIAIRES|PERIODE|SERVICE|ANNEXE)\b", name(c.value)) for c in occupied.values())
            if title_cue and neighbour_data:
                classifications[row] = RowClassification("TITLE", "Libellé de section fusionné entre des lignes individuelles.")
            else:
                classifications[row] = RowClassification("UNKNOWN", "Ligne fusionnée et peu remplie, sans identifiant : structure incertaine, à vérifier dans la classification.", confidence=.45)
            continue
        if styled and sparse and neighbour_data and not any(col in row_schema.columns.values() for col in occupied):
            classifications[row] = RowClassification("UNKNOWN", "Présentation distincte et clairsemée hors des champs d'identification, entre des enregistrements.", confidence=.4)
            continue
        classifications[row] = RowClassification("DATA", "Ligne du tableau sans indices structurels suffisants ; même partielle, elle reste contrôlée.")
    return classifications


REGION_TYPES = {"DATA": "DATA_REGION", "HEADER": "HEADER_REGION", "TITLE": "TITLE_REGION",
                "REPORT": "REPORT_REGION", "TOTAL": "SUMMARY_REGION", "SUBTOTAL": "SUMMARY_REGION",
                "SUMMARY": "SUMMARY_REGION", "NOTES": "NOTES_REGION", "EMPTY": "EMPTY_REGION",
                "OTHER": "UNKNOWN_REGION", "UNKNOWN": "UNKNOWN_REGION"}


def build_regions(classifications, schemas):
    regions = []
    for row, info in classifications.items():
        schema = schemas[row].header
        kind = REGION_TYPES[info.kind]
        if regions and regions[-1]["Type"] == kind and regions[-1]["En-têtes"] == schema and regions[-1]["Fin"] == row - 1:
            regions[-1]["Fin"] = row
            regions[-1]["Confiance"] = min(regions[-1]["Confiance"], info.confidence)
        else:
            regions.append({"Début": row, "Fin": row, "Type": kind, "En-têtes": schema,
                            "Confiance": info.confidence, "Raison": info.reason})
    return regions

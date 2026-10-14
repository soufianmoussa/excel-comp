"""Explicit column equivalences and cell-by-cell comparison of matched records."""
from dataclasses import dataclass, replace

import pandas as pd
from openpyxl.utils import get_column_letter

from comparator import STATUSES
from data_quality import FINDING_COLUMNS, default_rule, finding, infer_kind, parse_date, parse_number, semantic
from excel_reader import header_key
from normalizer import FIELDS, name, normalize, text


DIFF_COLUMNS = ["Champ", "Feuille A", "Ligne A", "Cellules A", "Valeur A", "Feuille B", "Ligne B", "Cellules B", "Valeur B", "Type", "Gravité", "Explication"]
MAPPING_COLUMNS = ["Actif", "Champ", "Source A", "Source B", "Mode", "Origine"]


@dataclass
class Source:
    columns: list
    role: str
    derived: bool = False


def sources(table):
    result = {f"{get_column_letter(col)} · {title}": Source([col], semantic(title))
              for col, title in table.all_columns.items()}
    for role, components in (("name", ["first", "last"]), ("rib", ["bank", "city", "account", "key"])):
        if role not in table.columns and all(c in table.columns for c in components):
            result[f"{FIELDS[role]} (reconstitué)"] = Source([table.columns[c] for c in components], role, True)
    return result


def detect_mappings(a, b, fields):
    sa, sb = sources(a), sources(b)
    rows, used_a, used_b = [], set(), set()
    for role in ("name", "ddp", "cin", "rib", "start", "end"):
        left = [k for k, s in sa.items() if s.role == role]
        right = [k for k, s in sb.items() if s.role == role]
        if len(left) == len(right) == 1:
            active = role not in FIELDS or role in fields
            rows.append([active, FIELDS.get(role, {"start": "Date de début", "end": "Date de fin"}.get(role)), left[0], right[0], role if role in FIELDS else "date", "Équivalence connue"])
            used_a.update(sa[left[0]].columns)
            used_b.update(sb[right[0]].columns)
    # Exact headers only: no fuzzy string similarity, and no guessed net-amount equivalence.
    for ka, source_a in sa.items():
        if source_a.derived or set(source_a.columns) & used_a:
            continue
        col = source_a.columns[0]
        if sum(header_key(title) == header_key(a.all_columns[col]) for title in a.all_columns.values()) != 1:
            continue
        candidates = [kb for kb, source_b in sb.items() if not source_b.derived and not set(source_b.columns) & used_b
                      and header_key(a.all_columns[col]) == header_key(b.all_columns[source_b.columns[0]])]
        if len(candidates) == 1:
            kb = candidates[0]
            mode = "number" if source_a.role in {"amount", "amount_net"} else "identifier" if source_a.role in {"bank", "city", "account", "key"} else "text"
            if mode == "text" and infer_kind(a, col) == infer_kind(b, sb[kb].columns[0]):
                mode = infer_kind(a, col)
            active = source_a.role not in FIELDS or source_a.role in fields
            rows.append([active, a.all_columns[col], ka, kb, mode, "Intitulé identique"])
            used_a.update(source_a.columns)
            used_b.update(sb[kb].columns)
    left = [k for k, s in sa.items() if s.role == "amount_net" and not set(s.columns) & used_a]
    right = [k for k, s in sb.items() if s.role == "amount_net" and not set(s.columns) & used_b]
    if len(left) == len(right) == 1:
        rows.append([False, "Montant net", left[0], right[0], "number", "À confirmer : même valeur métier ?"])
    return pd.DataFrame(rows, columns=MAPPING_COLUMNS)


def validate_mappings(frame, a, b):
    sa, sb = sources(a), sources(b)
    for row in frame.to_dict("records"):
        if not row.get("Actif") or pd.isna(row.get("Actif")):
            continue
        if row["Source A"] not in sa or row["Source B"] not in sb:
            raise ValueError("Chaque correspondance active doit indiquer une colonne dans les deux fichiers.")
        if row["Mode"] not in {*FIELDS, "text", "number", "date", "identifier"}:
            raise ValueError("Choisissez le type de comparaison de chaque correspondance active.")
        if not isinstance(row["Champ"], str) or not row["Champ"].strip():
            raise ValueError("Donnez un nom à chaque correspondance active.")


def regional_source(table, row, source):
    schema = table.row_schemas.get(row, table)
    if schema.header == table.header:
        return source
    if source.role in schema.columns:
        return Source([schema.columns[source.role]], source.role)
    parts = {"name": ["first", "last"], "rib": ["bank", "city", "account", "key"]}.get(source.role)
    if parts and all(p in schema.columns for p in parts):
        return Source([schema.columns[p] for p in parts], source.role, True)
    original = header_key(table.all_columns[source.columns[0]]) if source.columns else ""
    candidates = [col for col, title in schema.all_columns.items() if header_key(title) == original]
    if len(candidates) != 1 and source.role:
        candidates = [col for col, title in schema.all_columns.items() if semantic(title) == source.role]
    return Source(candidates if len(candidates) == 1 else [], source.role)


def source_value(table, record, source):
    if not source.columns:
        return "", "", "", True
    cells = [table.cells[record.row, col] for col in source.columns]
    raw = " | ".join(c.display for c in cells)
    uncertain = any(c.role == "MERGED_PLACEHOLDER" or c.data_type in {"f", "e"} or
                    source.role in {"cin", "ddp", "rib", "bank", "city", "account", "key"} and
                    isinstance(c.value, (int, float)) and abs(c.value) >= 10**15 for c in cells)
    if source.derived:
        value = record.values[source.role]
        uncertain |= source.role in record.issues and bool(value)
    else:
        value = cells[0].display if source.role in {"cin", "ddp", "rib", "bank", "city", "account", "key"} else cells[0].value
    return value, raw, ", ".join(dict.fromkeys(c.coordinate for c in cells)), uncertain


def comparable(mode, value):
    if value is None or not text(value).strip():
        return ""
    if mode in FIELDS:
        return normalize(mode, value)
    if mode == "date":
        return parse_date(value)
    if mode == "number":
        return parse_number(value)
    if mode == "identifier":
        return normalize("ddp", value)
    return " ".join(text(value).split())


def compare_mapped(a, b, record_results, mappings, rules_a=None, rules_b=None):
    validate_mappings(mappings, a, b)
    sa, sb = sources(a), sources(b)
    ra, rb = {r.row: r for r in a.records}, {r.row: r for r in b.records}
    differences, markers = [], []
    updated = record_results.copy()
    active = mappings.loc[mappings["Actif"].fillna(False).astype(bool)]
    # Matching finds candidates; the active mappings alone decide field differences.
    if not updated.empty:
        updated["Champs différents"] = ""
        previous_diff = updated["Code statut"] == "DIFFERENT"
        updated.loc[previous_diff, "Code statut"] = "MATCH"
        updated.loc[previous_diff, "Statut"] = STATUSES["MATCH"]
        for field in FIELDS:
            if f"Contrôle {FIELDS[field]}" in updated:
                updated[f"Contrôle {FIELDS[field]}"] = "Non comparé"
        if active.empty:
            matched = updated["Code statut"].isin(["MATCH", "DIFFERENT"])
            updated.loc[matched, "Code statut"] = "INCOMPLETE"
            updated.loc[matched, "Statut"] = STATUSES["INCOMPLETE"]
            updated.loc[matched, "Détails"] = "Personne retrouvée, mais aucune correspondance de colonnes n'est active."
    regional_plans = {}
    for index, result in record_results.iterrows():
        if result["Code statut"] in {"AMBIGUOUS", "MISSING_IN_B"} or not result["Lignes B"] or "," in result["Lignes B"]:
            continue
        record_a, record_b = ra[int(result["Ligne A"])], rb[int(result["Lignes B"])]
        if not active.empty:
            updated.at[index, "Code statut"] = "MATCH"
            updated.at[index, "Statut"] = STATUSES["MATCH"]
            updated.at[index, "À vérifier"] = ""
        changed, blocked = [], []
        schema_a, schema_b = a.row_schemas.get(record_a.row, a), b.row_schemas.get(record_b.row, b)
        plan_key = (schema_a.header, schema_b.header)
        if plan_key not in regional_plans:
            comparisons, covered = [], set()
            for mapping in mappings.to_dict("records"):
                source_a = regional_source(a, record_a.row, sa[mapping["Source A"]]) if mapping["Source A"] in sa else Source([], "")
                source_b = regional_source(b, record_b.row, sb[mapping["Source B"]]) if mapping["Source B"] in sb else Source([], "")
                covered.add((tuple(source_a.columns), tuple(source_b.columns)))
                comparisons.append((mapping, source_a, source_b))
            if schema_a.header != a.header or schema_b.header != b.header:
                region_a = replace(a, header=schema_a.header, columns=schema_a.columns, all_columns=schema_a.all_columns,
                                   rows=[r for r in a.rows if a.row_schemas[r].header == schema_a.header])
                region_b = replace(b, header=schema_b.header, columns=schema_b.columns, all_columns=schema_b.all_columns,
                                   rows=[r for r in b.rows if b.row_schemas[r].header == schema_b.header])
                extra_a, extra_b = sources(region_a), sources(region_b)
                for mapping in detect_mappings(region_a, region_b, list(FIELDS)).to_dict("records"):
                    source_a, source_b = extra_a[mapping["Source A"]], extra_b[mapping["Source B"]]
                    if (tuple(source_a.columns), tuple(source_b.columns)) not in covered:
                        comparisons.append((mapping, source_a, source_b))
            regional_plans[plan_key] = comparisons
        comparisons = regional_plans[plan_key]
        for mapping, source_a, source_b in comparisons:
            if not mapping["Actif"] or pd.isna(mapping["Actif"]):
                continue
            va, raw_a, coords_a, uncertain_a = source_value(a, record_a, source_a)
            vb, raw_b, coords_b, uncertain_b = source_value(b, record_b, source_b)
            try:
                for table, rec, src, rules in ((a, record_a, source_a, rules_a or {}), (b, record_b, source_b, rules_b or {})):
                    for col in src.columns:
                        schema = table.row_schemas.get(rec.row, table)
                        rule = rules.get((table.sheet, col))
                        if rule is None or (schema.header != table.header and rule == default_rule(table.all_columns[col])):
                            rule = default_rule(schema.all_columns[col])
                        if rule.required and not table.cells[rec.row, col].display.strip():
                            raise ValueError("missing required comparison value")
                if mapping["Mode"] in {"identifier", "cin", "ddp", "rib"}:
                    for table, rec, src in ((a, record_a, source_a), (b, record_b, source_b)):
                        if any(isinstance(table.cells[rec.row, c].value, (int, float)) and
                               abs(table.cells[rec.row, c].value) >= 10**15 for c in src.columns):
                            raise ValueError("uncertain identifier")
                if uncertain_a or uncertain_b:
                    raise ValueError("uncertain source")
                equal = comparable(mapping["Mode"], va) == comparable(mapping["Mode"], vb)
                if equal:
                    if mapping["Mode"] in FIELDS:
                        updated.at[index, f"Contrôle {FIELDS[mapping['Mode']]}"] = "✓ Identique" if text(va).strip() else "? Incomplet"
                    continue
                category, severity = "Différence", "Erreur"
                message = f"Différence détectée : {mapping['Champ']}. Aucun fichier n'est présumé correct ; vérifiez les deux valeurs."
                changed.append(mapping["Champ"])
            except (ValueError, TypeError):
                category, severity = "Comparaison impossible", "Avertissement"
                message = f"{mapping['Champ']} : valeur incertaine ou non interprétable ; aucune différence n'est affirmée."
                blocked.append(mapping["Champ"])
            if mapping["Mode"] in FIELDS:
                updated.at[index, f"Contrôle {FIELDS[mapping['Mode']]}"] = "✗ Différent" if category == "Différence" else "? Incomplet"
            differences.append(dict(zip(DIFF_COLUMNS, [mapping["Champ"], a.sheet, record_a.row, coords_a, raw_a,
                b.sheet, record_b.row, coords_b, raw_b, category, severity, message])))
            for label, table, rec, source in (("A", a, record_a, source_a), ("B", b, record_b, source_b)):
                for col in source.columns:
                    markers.append(finding(label, table, rec.row, col, category, severity, message))
        if changed:
            existing = updated.at[index, "Champs différents"]
            updated.at[index, "Champs différents"] = ", ".join(dict.fromkeys(([existing] if existing else []) + changed))
            if updated.at[index, "Code statut"] == "MATCH":
                updated.at[index, "Code statut"] = "DIFFERENT"
                updated.at[index, "Statut"] = STATUSES["DIFFERENT"]
        if blocked:
            updated.at[index, "Code statut"] = "INCOMPLETE"
            updated.at[index, "Statut"] = STATUSES["INCOMPLETE"]
            updated.at[index, "À vérifier"] = ", ".join(blocked)
    return updated, pd.DataFrame(differences, columns=DIFF_COLUMNS), pd.DataFrame(markers, columns=FINDING_COLUMNS)


def unmatched_columns(tables_a, tables_b, a, b, mappings):
    used = {"A": set(), "B": set()}
    for row in mappings.to_dict("records"):
        if row["Actif"] and not pd.isna(row["Actif"]):
            used["A"].update(sources(a)[row["Source A"]].columns)
            used["B"].update(sources(b)[row["Source B"]].columns)
    rows = []
    for label, tables, main in (("A", tables_a, a), ("B", tables_b, b)):
        for table in tables:
            for col, title in table.all_columns.items():
                if table.sheet != main.sheet or col not in used[label]:
                    rows.append({"Fichier": label, "Feuille": table.sheet, "Colonne": get_column_letter(col),
                        "Titre": title, "Gravité": "Information", "Explication": "Contrôles internes uniquement ; aucune correspondance active."})
    return pd.DataFrame(rows, columns=["Fichier", "Feuille", "Colonne", "Titre", "Gravité", "Explication"])

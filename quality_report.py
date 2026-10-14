"""Orchestrate independent audits and expose non-overlapping reporting units."""
import pandas as pd

from comparator import compare, STATUSES
from data_quality import analyze_table, FINDING_COLUMNS
from field_mapping import compare_mapped, unmatched_columns
from normalizer import FIELDS


def missing_people(a, b, fields):
    """List absences in both directions without treating uncertainty as absence."""
    columns = ["Fichier", "Absent du fichier", "Feuille", "Ligne Excel", *FIELDS.values()]
    common = [f for f in fields if f in a.available & b.available]
    if not common:
        return pd.DataFrame(columns=columns)
    forward = compare(a.records, b.records, common)
    reverse = compare(b.records, a.records, common)

    def referenced(frame):
        return {int(row.strip()) for value in frame["Lignes B"] for row in str(value).split(",") if row.strip()}

    rows = []
    # A fallback by name can be asymmetric. Any candidate in either direction
    # prevents us from declaring the referenced record absent.
    for label, other, table, matches, candidates in (
            ("A", "B", a, forward, referenced(reverse)),
            ("B", "A", b, reverse, referenced(forward))):
        missing = set(matches.loc[matches["Code statut"] == "MISSING_IN_B", "Ligne A"]) - candidates
        for record in table.records:
            if record.row in missing:
                rows.append([label, other, table.sheet, record.row,
                             *(record.values.get(field, "") for field in FIELDS)])
    return pd.DataFrame(rows, columns=columns)


def errors_only_report(report):
    """Filter user-facing findings without turning uncertain matches into successes."""
    result = dict(report)
    for key in ("findings", "anomalies", "differences", "visible_findings"):
        if key in result:
            result[key] = result[key].loc[result[key]["Gravité"] == "Erreur"].copy()
    anomalies, findings = result["anomalies"], result["findings"]
    stats = dict(report["stats"])
    stats.update({
        "Anomalies du fichier A": int((anomalies["Fichier"] == "A").sum()),
        "Anomalies du fichier B": int((anomalies["Fichier"] == "B").sum()),
        "Différences entre fichiers": len(result["differences"]),
        "Cellules vides problématiques": int((anomalies["Type"] == "Valeur manquante").sum()),
        "Erreurs de format": int(anomalies["Type"].isin(["Type", "Longueur", "Format RIB", "Format identifiant"]).sum()),
        "Incohérences de dates": int(anomalies["Type"].isin(["Date invalide", "Plage de dates", "Chronologie"]).sum()),
        "Cellules signalées (distinctes)": len(findings[["Fichier", "Feuille", "Cellule"]].drop_duplicates()),
    })
    stats.pop("Comparaisons impossibles", None)
    result["stats"] = stats
    result["unmatched"] = result["unmatched"].drop(columns="Gravité", errors="ignore")
    return result


def presentation_summary(report, tables_a, tables_b, a, b):
    """Count confirmed errors only; matching uncertainty is not an error."""
    items = report["findings"].loc[report["findings"]["Gravité"] == "Erreur"].copy()
    people = {(label, t.sheet, r.row): r.values.get("name", "") for label, tables in (("A", tables_a), ("B", tables_b)) for t in tables for r in t.records}
    items["Personne"] = [people.get((r["Fichier"], r["Feuille"], r["Ligne"]), "") for r in items.to_dict("records")]
    all_rows = {(label, t.sheet, row) for label, tables in (("A", tables_a), ("B", tables_b)) for t in tables for row in t.rows}
    affected = set(items[["Fichier", "Feuille", "Ligne"]].itertuples(index=False, name=None))
    return items, {"Nombre de lignes analysées": len(all_rows), "Nombre d’erreurs détectées": len(items),
                   "Nombre de lignes sans erreur détectée": len(all_rows - affected)}


def analyze(tables_a, tables_b, a, b, fields, mappings, rules_a, rules_b):
    internal = []
    for label, tables, rules in (("A", tables_a, rules_a), ("B", tables_b, rules_b)):
        for table in tables:
            selected = {col: rule for (sheet, col), rule in rules.items() if sheet == table.sheet}
            internal.append(analyze_table(table, label, selected))
    anomalies = pd.concat(internal, ignore_index=True) if internal else pd.DataFrame(columns=FINDING_COLUMNS)
    common_fields = [f for f in fields if f in a.available & b.available]
    records = compare(a.records, b.records if common_fields else [], common_fields or fields or ["cin"])
    if not common_fields and not records.empty:
        records["Code statut"] = "INCOMPLETE"
        records["Statut"] = STATUSES["INCOMPLETE"]
        records["Détails"] = "Aucun identifiant commun sélectionné : analyse interne uniquement."
    records, differences, markers = compare_mapped(a, b, records, mappings, rules_a, rules_b)
    unmatched = unmatched_columns(tables_a, tables_b, a, b, mappings)
    all_findings = pd.concat([anomalies, markers], ignore_index=True).drop_duplicates()
    stats = {
        "Total de colonnes analysées": sum(len(t.all_columns) for t in [*tables_a, *tables_b]),
        "Total de cellules analysées": sum(len(t.rows) * len(t.all_columns) for t in [*tables_a, *tables_b]),
        "Cellules classifiées (structure)": sum(len(t.cells) for t in [*tables_a, *tables_b]),
        "Lignes hors données conservées": sum(len(t.original_rows) - len(t.rows) for t in [*tables_a, *tables_b]),
        "Montants de synthèse examinés": sum(len(info.amount_columns) for t in [*tables_a, *tables_b] for info in t.row_classifications.values()),
        "Anomalies du fichier A": int((anomalies["Fichier"] == "A").sum()),
        "Anomalies du fichier B": int((anomalies["Fichier"] == "B").sum()),
        "Différences entre fichiers": int((differences["Type"] == "Différence").sum()),
        "Cellules vides problématiques": int((anomalies["Type"] == "Valeur manquante").sum()),
        "Erreurs de format": int((anomalies["Type"].isin(["Type", "Longueur", "Format RIB", "Format identifiant"]) & (anomalies["Gravité"] == "Erreur")).sum()),
        "Incohérences de dates": int(anomalies["Type"].isin(["Date invalide", "Plage de dates", "Chronologie"]).sum()),
        "Colonnes sans correspondance": len(unmatched),
        "Cellules signalées (distinctes)": len(all_findings[["Fichier", "Feuille", "Cellule"]].drop_duplicates()),
        "Comparaisons impossibles": int((differences["Type"] == "Comparaison impossible").sum()),
    }
    return {"records": records, "anomalies": anomalies, "differences": differences,
            "findings": all_findings, "unmatched": unmatched, "stats": stats,
            "missing_people": missing_people(a, b, fields)}

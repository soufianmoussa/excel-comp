"""Conservative matching by selected identifiers, independent of row order."""
from collections import defaultdict

import pandas as pd

from normalizer import FIELDS, normalize


STATUSES = {
    "MATCH": "✓ Correspondance", "MISSING_IN_B": "✗ Absent",
    "DIFFERENT": "⚠ Différence", "INCOMPLETE": "? Incomplet", "AMBIGUOUS": "? Correspondance ambiguë",
}


def compare(records_a, records_b, fields):
    fields = list(dict.fromkeys(fields))
    if not fields or any(f not in FIELDS for f in fields):
        raise ValueError("Choisissez au moins un champ à comparer.")
    def normalized(record):
        return {f: normalize(f, record.values.get(f)) for f in fields}
    aa, bb = [normalized(r) for r in records_a], [normalized(r) for r in records_b]
    indexes = {f: defaultdict(set) for f in fields}
    for j, (record, values) in enumerate(zip(records_b, bb)):
        for f, value in values.items():
            if value and f not in record.issues:
                indexes[f][value].add(j)
    candidates = []
    for record, values in zip(records_a, aa):
        strong = set()
        for f in fields:
            if f != "name" and values[f] and f not in record.issues:
                strong.update(indexes[f].get(values[f], set()))
        if not strong and "name" in fields and "name" not in record.issues and values["name"]:
            strong.update(indexes["name"].get(values["name"], set()))
        candidates.append(strong)
    rows = []
    for record, values, matches in zip(records_a, aa, candidates):
        b = records_b[next(iter(matches))] if len(matches) == 1 else None
        bv = normalized(b) if b else {}
        missing = [f for f in fields if not values[f] or f in record.issues or
                   (b is not None and (not bv[f] or f in b.issues))]
        different = [f for f in fields if b is not None and f not in missing and values[f] != bv[f]]
        ambiguous = len(matches) > 1
        if ambiguous:
            status = "AMBIGUOUS"
        elif missing:
            status = "INCOMPLETE"
        elif not matches:
            status = "MISSING_IN_B"
        elif different:
            status = "DIFFERENT"
        else:
            status = "MATCH"
        details = []
        for label, rec in (("A", record), ("B", b)):
            if rec:
                for f in fields:
                    if f in rec.issues:
                        details.append(f"Fichier {label}, {FIELDS[f]} : {rec.issues[f]}")
                    elif not normalize(f, rec.values.get(f)):
                        details.append(f"Fichier {label}, {FIELDS[f]} : information manquante.")
        if ambiguous:
            details.append("Plusieurs lignes B peuvent correspondre. Vérification nécessaire pour choisir la bonne personne.")
        elif not matches:
            details.append("Aucune correspondance trouvée avec les champs sélectionnés.")
        result = {"Code statut": status, "Statut": STATUSES[status], "Ligne A": record.row,
                  "Lignes B": ", ".join(str(records_b[j].row) for j in sorted(matches)),
                  "Champs différents": ", ".join(FIELDS[f] for f in different),
                  "À vérifier": ", ".join(FIELDS[f] for f in missing), "Détails": " ".join(details)}
        for f in FIELDS:
            result[f"{FIELDS[f]} fichier A"] = record.values.get(f, "")
            result[f"{FIELDS[f]} fichier B"] = b.values.get(f, "") if b else ""
            if f in fields:
                result[f"Contrôle {FIELDS[f]}"] = ("? Incomplet" if f in missing else
                    "⚠ Ambigu" if ambiguous else "— Absent" if b is None else
                    "✗ Différent" if f in different else "✓ Identique")
        rows.append(result)
    columns = ["Code statut", "Statut", "Ligne A", "Lignes B", "Champs différents", "À vérifier", "Détails"]
    for f in FIELDS:
        columns.extend([f"{FIELDS[f]} fichier A", f"{FIELDS[f]} fichier B"])
        if f in fields:
            columns.append(f"Contrôle {FIELDS[f]}")
    return pd.DataFrame(rows, columns=columns)

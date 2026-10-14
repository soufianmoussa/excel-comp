"""Local, French-language Streamlit application."""
import hashlib
import logging

import streamlit as st

import pandas as pd

from excel_reader import WorkbookError, detect_header, discover_tables, open_workbook, read_table
from exporter import SHEETS, export_quality
from normalizer import FIELDS
from quality_ui import configure, viewer, missing_people_view
from quality_report import analyze, presentation_summary, errors_only_report


st.set_page_config(page_title="Comparateur de fichiers Excel", page_icon="📊", layout="wide")
# Do not forward exception contents or source values to application logs.
logging.getLogger("openpyxl").setLevel(logging.ERROR)
st.title("Comparateur de fichiers Excel")
st.caption("Comparez les personnes et vérifiez la qualité de toutes les colonnes de vos fichiers Excel.")
st.info("Vos fichiers sont traités localement en mémoire. Aucune donnée n’est envoyée à un service externe.")


def upload_table(label):
    st.subheader(f"Fichier {label}")
    uploaded = st.file_uploader(f"Choisir le fichier {label}", type=["xlsx"], key=f"upload_{label}")
    if uploaded is None:
        return None, None
    data = uploaded.getvalue()
    digest = hashlib.sha256(data).hexdigest()
    st.success("✓ Fichier chargé")
    try:
        cache_key = f"upload_cache_{label}"
        cache = st.session_state.get(cache_key)
        if cache is None or cache["digest"] != digest or cache.get("version") != "structure-v1":
            workbook = open_workbook(data)
            detected = discover_tables(workbook)
            cache = {"version": "structure-v1", "digest": digest, "workbook": workbook, "detected": detected, "loaded": {}}
            st.session_state[cache_key] = cache
        workbook, detected = cache["workbook"], cache["detected"]
        if not detected:
            st.warning("Le fichier semble vide.")
            return None, None
        defaults = pd.DataFrame([{"Analyser": True, "Feuille": t.sheet, "Ligne des en-têtes": t.header,
                                  "Détection": t.detection} for t in detected])
        with st.expander("Options avancées"):
            st.caption("Toutes les feuilles non vides sont proposées. Vérifiez les en-têtes des feuilles inhabituelles. Vous pouvez corriger leur numéro de ligne ou exclure une feuille de référence.")
            sheet_config = st.data_editor(defaults, hide_index=True, key=f"headers_{label}_{digest}",
                disabled=["Feuille", "Détection"],
                column_config={"Ligne des en-têtes": st.column_config.NumberColumn(min_value=1, step=1, required=True)})
            st.caption("Correction facultative des zones : indiquez les lignes Excel de début et de fin. Toute correction nécessite de relancer l’analyse. DATA = données, HEADER = en-têtes, TITLE = titre, REPORT = report, TOTAL / SUBTOTAL / SUMMARY = synthèse, NOTES = notes, EMPTY = vide, UNKNOWN = incertain.")
            zones = st.data_editor(pd.DataFrame(columns=["Feuille", "Début", "Fin", "Type"]),
                key=f"zones_{label}_{digest}", num_rows="dynamic", hide_index=True,
                column_config={"Feuille": st.column_config.SelectboxColumn(options=workbook.sheetnames, required=True),
                    "Début": st.column_config.NumberColumn(min_value=1, step=1, required=True),
                    "Fin": st.column_config.NumberColumn(min_value=1, step=1, required=True),
                    "Type": st.column_config.SelectboxColumn(options=["DATA", "HEADER", "TITLE", "REPORT", "TOTAL", "SUBTOTAL", "SUMMARY", "NOTES", "EMPTY", "UNKNOWN"], required=True)})
        overrides = {}
        for zone in zones.to_dict("records"):
            if any(pd.isna(value) or value == "" for value in zone.values()):
                raise WorkbookError("Complétez la feuille, les deux lignes et le type de chaque correction de structure.")
            sheet_name, start, end = zone["Feuille"], int(zone["Début"]), int(zone["Fin"])
            if sheet_name not in workbook.sheetnames or start != float(zone["Début"]) or end != float(zone["Fin"]) or not 1 <= start <= end <= workbook[sheet_name].max_row:
                raise WorkbookError("La plage de correction de structure est invalide.")
            selected = overrides.setdefault(sheet_name, {})
            if any(row in selected for row in range(start, end + 1)):
                raise WorkbookError("Les corrections de structure ne doivent pas se chevaucher.")
            selected.update({row: zone["Type"] for row in range(start, end + 1)})
        config_key = sheet_config.to_json() + zones.to_json()
        if config_key not in cache["loaded"]:
            tables = []
            for item in sheet_config.to_dict("records"):
                if not item["Analyser"]:
                    continue
                row = int(item["Ligne des en-têtes"])
                sheet = workbook[item["Feuille"]]
                if not 1 <= row <= sheet.max_row:
                    raise WorkbookError(f"La ligne des en-têtes n'existe pas dans la feuille {sheet.title}.")
                table = detect_header(sheet, row)
                table.detection = item["Détection"]
                tables.append(read_table(workbook, table, overrides.get(table.sheet)))
            # Cache only the current header selection, inside this private session.
            cache["loaded"] = {config_key: tables}
        tables = cache["loaded"][config_key]
        populated = [t for t in tables if t.original_rows]
        if not populated:
            st.warning("Le fichier semble vide : aucune ligne de données détectée.")
            return None, None
        preferred = sorted(populated, key=lambda t: (not bool(t.records), -len(t.available)))
        chosen = preferred[0]
        if len(preferred) > 1:
            with st.expander("Feuille de rapprochement (facultatif)"):
                selected = st.selectbox("Feuille à rapprocher avec l'autre fichier", [t.sheet for t in preferred], key=f"sheet_{label}_{digest}")
            chosen = next(t for t in preferred if t.sheet == selected)
        st.success(f"✓ Tableau détecté · {chosen.sheet} · en-têtes ligne {chosen.header}\n\n✓ {len(chosen.records)} lignes détectées")
        st.caption(f"Analyse interne : {len(tables)} feuille(s), {sum(len(t.all_columns) for t in tables)} colonnes. Rapprochement des personnes sur la feuille choisie.")
        if any("proposés" in t.detection for t in tables):
            st.warning("Certaines feuilles ont des en-têtes proposés automatiquement : vérifiez-les dans Options avancées.")
        st.session_state[f"tables_{label}"] = tables
        return chosen, (digest, chosen.sheet, config_key)
    except WorkbookError as error:
        st.error(str(error))
    except Exception:
        st.error(f"Impossible de lire le fichier {label}. Vérifiez son contenu et réenregistrez-le au format .xlsx.")
    return None, None


left, right = st.columns(2, gap="large")
with left:
    table_a, key_a = upload_table("A")
with right:
    table_b, key_b = upload_table("B")

fields = list(FIELDS)  # Internal record identifiers; every column is still audited.
ready = table_a is not None and table_b is not None
mapping, edited, parsed = None, {}, {}
if ready:
    config_id = hashlib.sha256(repr((key_a, key_b)).encode()).hexdigest()[:16]
    mapping, edited, parsed, valid = configure(st.session_state["tables_A"], st.session_state["tables_B"], table_a, table_b, fields, config_id)
    ready = ready and valid

signature = ("presence-v1", key_a, key_b, tuple(fields), None if mapping is None else mapping.to_json(),
             tuple((label, frame.to_json()) for label, frame in edited.items()))
if st.session_state.get("result_signature") != signature:
    for key in ("results", "export", "result_signature", "report"):
        st.session_state.pop(key, None)

if st.button("Analyser et comparer", type="primary", width="stretch", disabled=not ready):
    try:
        with st.spinner("Comparaison en cours…"):
            tables_a, tables_b = st.session_state["tables_A"], st.session_state["tables_B"]
            report = analyze(tables_a, tables_b, table_a, table_b, fields, mapping, parsed["A"], parsed["B"])
            report = errors_only_report(report)
            report["visible_findings"], report["main_metrics"] = presentation_summary(report, tables_a, tables_b, table_a, table_b)
            results = report["records"]
            st.session_state["export"] = export_quality(report, st.session_state["upload_cache_A"]["workbook"],
                st.session_state["upload_cache_B"]["workbook"], tables_a, tables_b, table_a, table_b, mapping, edited["A"], edited["B"])
            st.session_state["report"] = report
            st.session_state["results"] = results
            st.session_state["result_signature"] = signature
    except Exception:
        st.error("La comparaison n’a pas pu aboutir. Vérifiez les deux fichiers et réessayez.")

if "results" in st.session_state:
    results = st.session_state["results"]
    report = st.session_state["report"]
    st.subheader("Résultats")
    for col, (label, value) in zip(st.columns(3), report["main_metrics"].items()):
        col.metric(label, value)
    st.caption("Seules les erreurs sont affichées et comptées. Les compteurs de lignes concernent les données individuelles des deux fichiers, répétitions comprises. Les titres, reports et totaux restent consultables. Une ligne peut contenir plusieurs erreurs ; l’absence d’erreur détectée ne garantit pas qu’un rapprochement a pu être effectué.")
    missing_people_view(report, table_a, table_b)
    viewer(report, st.session_state["tables_A"], st.session_state["tables_B"], table_a, table_b)
    with st.expander("Détails de l’analyse"):
        tabs = st.tabs(["Tous", *SHEETS.values()])
        for tab, status in zip(tabs, [None, *SHEETS]):
            with tab:
                frame = results if status is None else results.loc[results["Code statut"] == status]
                st.dataframe(frame.drop(columns="Code statut"), hide_index=True, width="stretch")
        st.markdown("**Erreurs détaillées**")
        st.dataframe(report["anomalies"], hide_index=True, width="stretch")
        st.markdown("**Différences entre fichiers — valeurs et coordonnées**")
        st.dataframe(report["differences"], hide_index=True, width="stretch")
        st.markdown("**Colonnes sans correspondance**")
        st.dataframe(report["unmatched"], hide_index=True, width="stretch")
        st.dataframe(pd.DataFrame(report["stats"].items(), columns=["Indicateur", "Valeur"]), hide_index=True, width="stretch")
    st.download_button("Télécharger le rapport Excel", data=st.session_state["export"], file_name="resultat_comparaison.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", width="stretch")
    st.caption("Le rapport inclut les feuilles annotées, leurs coordonnées, les règles appliquées et les correspondances. Les formules sources sont conservées comme texte non exécutable ; vos fichiers d'origine restent intacts.")

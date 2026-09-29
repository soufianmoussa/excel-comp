"""Local, French-language Streamlit application."""
import hashlib
import logging

import streamlit as st

from comparator import STATUSES, compare
from excel_reader import WorkbookError, detect_header, detect_tables, open_workbook, read_table
from exporter import SHEETS, export_excel
from normalizer import FIELDS


st.set_page_config(page_title="Comparateur de fichiers Excel", page_icon="📊", layout="wide")
# Do not forward exception contents or source values to application logs.
logging.getLogger("openpyxl").setLevel(logging.ERROR)
st.title("Comparateur de fichiers Excel")
st.caption("Comparez automatiquement les informations entre deux fichiers Excel.")
st.info("Vos fichiers sont traités localement en mémoire. Aucune donnée n’est envoyée à un service externe.")


def upload_table(label):
    st.subheader(f"Fichier {label}")
    uploaded = st.file_uploader(f"Choisir le fichier {label}", type=["xlsx"], key=f"upload_{label}")
    if uploaded is None:
        return None, None
    data = uploaded.getvalue()
    digest = hashlib.sha256(data).hexdigest()
    st.success("✓ Fichier chargé")
    workbook = None
    try:
        workbook = open_workbook(data)
        tables = detect_tables(workbook)
        chosen = None
        if len(tables) == 1:
            chosen = tables[0]
        elif tables:
            sheet = st.selectbox("Plusieurs tableaux détectés : choisissez la feuille", [t.sheet for t in tables], key=f"sheet_{label}_{digest}")
            chosen = next(t for t in tables if t.sheet == sheet)
        else:
            st.warning(f"Le tableau principal n’a pas pu être détecté dans le fichier {label}. Utilisez les options avancées.")
        with st.expander("Options avancées"):
            manual = st.checkbox("Choisir la feuille et la ligne des en-têtes", key=f"manual_{label}_{digest}")
            if manual:
                sheet = st.selectbox("Feuille", workbook.sheetnames, key=f"manual_sheet_{label}_{digest}")
                row = st.number_input("Ligne des en-têtes (la première ligne vaut 1)", min_value=1, max_value=max(1, workbook[sheet].max_row), value=1, step=1, key=f"row_{label}_{digest}_{sheet}")
                chosen = detect_header(workbook[sheet], int(row))
                if not chosen.available:
                    st.warning("Aucune colonne reconnue sur cette ligne. Vérifiez la feuille et la ligne choisies.")
                    return None, None
            st.caption("Les deux présentations sont reconnues : nom complet et RIB complet, ou nom/prénom et RIB séparés.")
        if chosen is None:
            return None, None
        table = read_table(workbook, chosen)
        if not table.records:
            st.warning("Le fichier semble vide : aucune ligne de données détectée.")
            return None, None
        st.success(f"✓ Tableau détecté · {table.sheet} · en-têtes ligne {table.header}\n\n✓ {len(table.records)} lignes détectées")
        uncertain = sum(bool(r.issues) for r in table.records)
        if uncertain:
            st.warning(f"{uncertain} ligne(s) contiennent des informations manquantes ou à vérifier. Le résultat précisera les champs concernés.")
        return table, (digest, table.sheet, table.header, table.end_header)
    except WorkbookError as error:
        st.error(str(error))
    except Exception:
        st.error(f"Impossible de lire le fichier {label}. Vérifiez son contenu et réenregistrez-le au format .xlsx.")
    finally:
        if workbook is not None:
            workbook.close()
    return None, None


left, right = st.columns(2, gap="large")
with left:
    table_a, key_a = upload_table("A")
with right:
    table_b, key_b = upload_table("B")

st.subheader("Données à comparer")
fields = []
for col, (field, title) in zip(st.columns(4), FIELDS.items()):
    with col:
        if st.checkbox(title, value=True, key=f"field_{field}"):
            fields.append(field)
st.caption("Les lignes sont rapprochées grâce aux identifiants sélectionnés ; le nom sert de recours. L’ordre des lignes n’a pas d’importance.")
if not fields:
    st.warning("Choisissez au moins un champ à comparer.")
ready = table_a is not None and table_b is not None and bool(fields)
if ready:
    for label, table in (("A", table_a), ("B", table_b)):
        for field in fields:
            if field not in table.available:
                st.error(f"Impossible de trouver la colonne {FIELDS[field]} dans le fichier {label}. Vérifiez les en-têtes ou décochez ce champ.")
                ready = False

signature = (key_a, key_b, tuple(fields))
if st.session_state.get("result_signature") != signature:
    for key in ("results", "export", "result_signature"):
        st.session_state.pop(key, None)

if st.button("Comparer les fichiers", type="primary", use_container_width=True, disabled=not ready):
    try:
        with st.spinner("Comparaison en cours…"):
            results = compare(table_a.records, table_b.records, fields)
            metadata = {"Feuille A": table_a.sheet, "En-têtes A": table_a.header,
                        "Feuille B": table_b.sheet, "En-têtes B": table_b.header,
                        "Champs comparés": ", ".join(FIELDS[f] for f in fields),
                        "Périmètre": "Toutes les lignes A ; les lignes présentes uniquement dans B ne sont pas comptées."}
            st.session_state["export"] = export_excel(results, metadata)
            st.session_state["results"] = results
            st.session_state["result_signature"] = signature
    except Exception:
        st.error("La comparaison n’a pas pu aboutir. Vérifiez les deux fichiers et réessayez.")

if "results" in st.session_state:
    results = st.session_state["results"]
    st.subheader("Résultat de la comparaison")
    metrics = st.columns(6)
    metrics[0].metric("Total analysé", len(results))
    for col, status in zip(metrics[1:], STATUSES):
        col.metric({"MATCH": "Correspondances", "MISSING_IN_B": "Absents du fichier B", "DIFFERENT": "Différences", "INCOMPLETE": "Données incomplètes", "DUPLICATE": "Doublons"}[status], int((results["Code statut"] == status).sum()))
    st.caption("Une ligne de résultat par ligne du fichier A. Les lignes présentes uniquement dans B ne sont pas comptées. Les données incertaines restent à vérifier.")
    tabs = st.tabs(["Tous", *SHEETS.values()])
    for tab, status in zip(tabs, [None, *SHEETS]):
        with tab:
            frame = results if status is None else results.loc[results["Code statut"] == status]
            if frame.empty:
                st.info("Aucune ligne dans cette catégorie.")
            else:
                st.dataframe(frame.drop(columns="Code statut"), hide_index=True, use_container_width=True)
    st.download_button("Télécharger le résultat Excel", data=st.session_state["export"], file_name="resultat_comparaison.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)

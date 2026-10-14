from io import BytesIO
from pathlib import Path

import streamlit as st
from openpyxl import Workbook
from streamlit.testing.v1 import AppTest


APP = str(Path(__file__).resolve().parents[1] / "app.py")


def workbook_data(day):
    w = Workbook()
    s = w.active
    s.append(["CIN", "DDR", "DEBUT", "Note"])
    s.append(["X1", 1, day, " Texte "])
    s.append(["X2", 2, "31/02/2025", None])
    data = BytesIO()
    w.save(data)
    return data.getvalue()


def test_viewer_filters_navigation_and_configuration_changes(monkeypatch):
    data = {"upload_A": workbook_data("01/10/2025"), "upload_B": workbook_data("02/10/2025")}
    monkeypatch.setattr(st, "file_uploader", lambda *args, **kwargs: BytesIO(data[kwargs["key"]]))
    app = AppTest.from_file(APP).run(timeout=15)
    assert not app.exception
    app.button[0].click().run(timeout=15)
    assert not app.exception and not app.error
    assert app.session_state["report"]["stats"]["Différences entre fichiers"] == 1
    assert len(app.metric) == 3
    assert not any(str(c.key).startswith("field_") for c in app.checkbox)
    assert not any(s.value == "Données à comparer" for s in app.subheader)
    assert app.button[0].label == "Analyser et comparer"
    assert len(app.get("download_button")) == 1
    assert app.get("download_button")[0].label == "Télécharger le rapport Excel"
    nav = next(widget for widget in app.selectbox if widget.label == "Naviguer entre les anomalies")
    previous = nav.value
    next(button for button in app.button if button.label == "Suivante →").click().run()
    nav = next(widget for widget in app.selectbox if widget.label == "Naviguer entre les anomalies")
    assert nav.value != previous
    assert not any(widget.label == "Gravité" for widget in app.multiselect)
    assert not app.exception
    next(widget for widget in app.selectbox if widget.label == "Fichier").select("B").run()
    assert not app.exception
    app.text_input[0].set_value("no-such-value").run()
    assert not app.exception
    assert any("Aucune anomalie" in info.value for info in app.info)
    # Change a real data-editor rule: make Note required in A; stale reports disappear.
    key = next(k for k in app.session_state.filtered_state if k.startswith("rules_A_"))
    app.session_state[key] = {"edited_rows": {3: {"required": True}}, "added_rows": [], "deleted_rows": []}
    app.run()
    assert "report" not in app.session_state
    # AppTest has no data_editor input API and replays the old widget protobuf;
    # inject the same browser delta for the subsequent button event as well.
    app.session_state[key] = {"edited_rows": {3: {"required": True}}, "added_rows": [], "deleted_rows": []}
    app.button[0].click().run(timeout=15)
    findings = app.session_state["report"]["anomalies"]
    assert len(findings.query("Fichier == 'A' and Titre == 'Note' and Type == 'Valeur manquante'")) == 1
    assert not app.exception



def test_click_anomaly_opens_its_original_cell(monkeypatch):
    from types import SimpleNamespace
    data = workbook_data("01/10/2025")
    monkeypatch.setattr(st, "file_uploader", lambda *args, **kwargs: BytesIO(data))
    original = st.dataframe
    clicked = {"rows": []}
    def dataframe(*args, **kwargs):
        event = original(*args, **kwargs)
        if str(kwargs.get("key", "")).startswith("list_anomaly_"):
            return SimpleNamespace(selection=SimpleNamespace(rows=clicked["rows"]))
        return event
    monkeypatch.setattr(st, "dataframe", dataframe)
    app = AppTest.from_file(APP).run()
    app.button[0].click().run(timeout=15)
    findings = app.session_state["report"]["visible_findings"].reset_index(drop=True)
    clicked["rows"] = [len(findings) - 1]
    app.run()
    assert not app.exception
    expected = findings.iloc[-1]
    nav = next(w for w in app.selectbox if w.label == "Naviguer entre les anomalies")
    assert nav.value == len(findings) - 1
    assert any(f"fichier {expected['Fichier']}, feuille {expected['Feuille']}, cellule {expected['Cellule']}" in info.value for info in app.info)


def test_invalid_admin_rule_blocks_comparison_with_french_message(monkeypatch):
    data = workbook_data("01/10/2025")
    monkeypatch.setattr(st, "file_uploader", lambda *args, **kwargs: BytesIO(data))
    app = AppTest.from_file(APP).run()
    key = next(k for k in app.session_state.filtered_state if k.startswith("rules_A_"))
    app.session_state[key] = {"edited_rows": {3: {"minimum": "invalide"}}, "added_rows": [], "deleted_rows": []}
    app.run()
    assert not app.exception
    assert app.button[0].disabled
    assert any("configuration" in error.value for error in app.error)


def test_generic_workbooks_still_analyze_without_identifiers(monkeypatch):
    w = Workbook()
    w.active.append(["Commentaire", "Mesure"])
    w.active.append([" suivi ", 2])
    data = BytesIO()
    w.save(data)
    monkeypatch.setattr(st, "file_uploader", lambda *args, **kwargs: BytesIO(data.getvalue()))
    app = AppTest.from_file(APP).run()
    app.button[0].click().run(timeout=15)
    assert not app.exception and not app.error
    report = app.session_state["report"]
    assert report["stats"]["Total de colonnes analysées"] == 4
    assert report["anomalies"].empty
    assert report["records"].iloc[0]["Code statut"] == "INCOMPLETE"


def test_current_samples_end_to_end_without_persisting_personal_data(monkeypatch):
    root = Path(APP).parent
    paths = {"upload_A": root / "FILE_A.xlsx", "upload_B": root / "File_B.xlsx"}
    if not all(p.exists() for p in paths.values()):
        import pytest
        pytest.skip("Exemples locaux absents")
    data = {k: p.read_bytes() for k, p in paths.items()}
    monkeypatch.setattr(st, "file_uploader", lambda *args, **kwargs: BytesIO(data[kwargs["key"]]))
    app = AppTest.from_file(APP).run(timeout=20)
    assert not app.exception
    app.button[0].click().run(timeout=20)
    assert not app.exception and not app.error
    assert app.session_state["report"]["stats"]["Total de colonnes analysées"] > 8
    assert len(app.session_state["export"]) > 0

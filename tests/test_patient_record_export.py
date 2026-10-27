# type: ignore
"""Tests für den Export "Patientenakte" (reine Logik, ohne GUI)."""
import datetime
import os
import shutil
import sys
import tempfile
import unittest

import pandas as pd

current_dir = os.path.dirname(os.path.abspath(__file__))
root_dir = os.path.abspath(os.path.join(current_dir, ".."))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from src.patient_record import patient_record_export as pr


def write_csv(folder, name, rows, encoding="utf-8"):
    """Schreibt eine CSV wie der Datenbank-Export: Semikolon, Werte in Anführungszeichen, NULL für leer."""
    header = list(rows[0].keys())
    lines = [";".join(header)]
    for r in rows:
        lines.append(";".join("NULL" if r[h] is None else f'"{r[h]}"' for h in header))
    with open(os.path.join(folder, f"{name}.csv"), "w", encoding=encoding) as f:
        f.write("\n".join(lines) + "\n")


class PatientRecordTestBase(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.mkdtemp()
        write_csv(self.folder, "patienten", [
            {"id": "P1", "p_nr": "P1", "p_anrede": "Frau", "p_name": " Öztürk ", "p_vname": "Ayşe",
             "p_plz": "01067", "p_ort": "Dresden", "p_street": "Hauptstr.", "p_hausnummer": "5a",
             "p_birth": "1950-03-01", "p_tel": "0351 123", "p_privatversichert": None,
             "p_zuzahlungsbefreit": "t", "rechnungsempfaenger": "null"},
            {"id": "P2", "p_nr": "P2", "p_anrede": "Herr", "p_name": "Adam", "p_vname": "Bernd",
             "p_plz": "23570", "p_ort": "Lübeck", "p_street": "Weg", "p_hausnummer": "1",
             "p_birth": "1890-12-24", "p_tel": None, "p_privatversichert": "t",
             "p_zuzahlungsbefreit": None,
             "rechnungsempfaenger": '[{""anrede"":""Herr"",""name"":""Sohn Adam"",""plz"":""23570"",'
                                    '""ort"":""Lübeck"",""strassepostfach"":""Weg 2""}]'},
        ])
        write_csv(self.folder, "resources", [{"id": "R1", "name": "Uwe Therapeut"}])
        write_csv(self.folder, "events", [
            {"id": "E1", "start": "2026-01-05 10:20:00", "end": "2026-01-05 10:40:00", "duration": "20",
             "title": "KG", "resource": "R1", "patient_id": "P1", "wahrgenommen": "1",
             "nicht_wahrgenommen": None, "abgesagt": None, "abgerechnet": "t", "hausbesuch": "f",
             "eventsequenz": "1/6", "dok_stichwort": "Befund", "dokumentation": "Mobilisation HWS",
             "description": None, "memo": None},
            {"id": "E2", "start": "2026-01-02 08:00:00", "end": "2026-01-02 08:20:00", "duration": "20",
             "title": "KG", "resource": "R1", "patient_id": "P1", "wahrgenommen": "0",
             "nicht_wahrgenommen": "1", "abgesagt": None, "abgerechnet": None, "hausbesuch": None,
             "eventsequenz": None, "dok_stichwort": None, "dokumentation": None,
             "description": None, "memo": None},
            {"id": "E3", "start": "2026-01-03 12:00:00", "end": "2026-01-03 13:00:00", "duration": "60",
             "title": "Flexzeit", "resource": "R1", "patient_id": None, "wahrgenommen": None,
             "nicht_wahrgenommen": None, "abgesagt": None, "abgerechnet": None, "hausbesuch": None,
             "eventsequenz": None, "dok_stichwort": None, "dokumentation": None,
             "description": "Flexzeit", "memo": None},
        ])
        write_csv(self.folder, "rezepte", [
            {"id": "Z1", "patienten_id": "P1", "rezeptnr": "152", "datum": "2025-12-20", "rezept_typ": "GK",
             "verordnungstyp": None, "krankenkasse": "AOK", "arztnr": "089585001",
             "vorrangiges_heilmittel": "KG", "behandlungseinheiten": "6",
             "vorrangiges_heilmittel2": "MT", "behandlungseinheiten2": "6",
             "leitsymptomatik_text": "Schädigung der Bewegungssegmente ; ", "hausbesuch": "Nein",
             "tp_therapiebericht": "x" * 40000},
        ])
        write_csv(self.folder, "rezepte_events", [
            {"id": "RE1", "patienten_id": "P1", "events_id": "E1", "rezepte_id": "Z1"},
            {"id": "RE2", "patienten_id": "P1", "events_id": "E2", "rezepte_id": "Z1"},
        ])
        write_csv(self.folder, "anatomy", [{"id": "A1", "name": "Female_Muscle_Front", "coordinates": "[]"}])
        write_csv(self.folder, "anatomy_patienten", [
            {"id": "AP1", "id_region": "A1", "name_region": None, "id_patient": "P1",
             "coordinate": "{}", "title": "Fraktur", "notes": "OP 2024", "timestamp": "2025-01-20"},
        ])
        write_csv(self.folder, "kartei", [{"id": "K1", "karteiname": "Anamnese ", "beschreibung": "x"}])
        write_csv(self.folder, "pat_property", [
            {"id": "Q1", "label": "Rauchen Sie?", "proptyp": "CHK", "sortierung": "1", "kartei_id": "K1"},
            {"id": "Q2", "label": "Schmerzstärke", "proptyp": "RDO", "sortierung": "0", "kartei_id": "K1"},
        ])
        write_csv(self.folder, "kartei_history", [
            {"id": "H1", "kartei_id": "K1", "patienten_id": "P2", "datum": "2025-04-09", "expires": None},
        ])
        write_csv(self.folder, "kartei_content", [
            {"id": "C1", "kartei_history_id": "H1", "kartei_property_id": "Q1", "content": "1"},
            {"id": "C2", "kartei_history_id": "H1", "kartei_property_id": "Q2", "content": "4"},
        ])
        write_csv(self.folder, "history", [
            {"id": "N1", "text": "Rückruf", "bearbeiter": None, "datum": "2025-10-01 08:39:00",
             "stichwort": "Telefon", "reftable": "patienten", "refid": "P2"},
            {"id": "N2", "text": "Praxisnotiz", "bearbeiter": None, "datum": "2025-10-01 08:39:00",
             "stichwort": "x", "reftable": "praxis", "refid": "X"},
        ])
        write_csv(self.folder, "files", [
            {"id": "F1", "reftable": "rezepte", "refid": "Z1", "size": "2048", "ext": "jpeg",
             "orgname": "Rezept.jpeg", "kategorie": "rezeptvorder", "erstellt_am": "2025-12-21",
             "filename": "F1-rezepte-Z1.jpeg", "beschreibung": None},
            {"id": "F2", "reftable": "praxis", "refid": "X", "size": "10", "ext": "png",
             "orgname": "logo.png", "kategorie": "logo", "erstellt_am": "2025-01-01",
             "filename": "logo.png", "beschreibung": None},
        ])

    def tearDown(self):
        shutil.rmtree(self.folder, ignore_errors=True)

    def sheets(self):
        tables, _ = pr.load_export_folder(self.folder)
        return pr.build_sheets(tables)


class TestHelpers(unittest.TestCase):
    def test_yes_no(self):
        s = pd.Series(["1", "0", "t", "f", "Ja", "nein", None, "?"])
        result = pr.yes_no(s).tolist()
        self.assertEqual(result[:6] + result[7:], ["Ja", "Nein", "Ja", "Nein", "Ja", "Nein", "?"])
        self.assertTrue(pd.isna(result[6]))

    def test_to_date(self):
        s = pd.Series(["2025-01-20", "2025-10-01 08:39:00.123", "1890-12-24", "kaputt", None])
        result = pr.to_date(s).tolist()
        self.assertEqual(result[0], datetime.date(2025, 1, 20))
        self.assertEqual(result[1], datetime.date(2025, 10, 1))
        self.assertEqual(result[2], "24.12.1890")  # vor 1900 -> Text
        self.assertEqual(result[3], "kaputt")
        self.assertTrue(pd.isna(result[4]))

    def test_view_label(self):
        self.assertEqual(pr.view_label("Male_Skeleton_Back"), "Skelett hinten (männlich)")
        self.assertEqual(pr.view_label("Unbekannt"), "Unbekannt")

    def test_sheet_name_sanitized_and_unique(self):
        used = set()
        self.assertEqual(pr.sheet_name("Fragebogen A/B", used), "Fragebogen A B")
        long = "Fragebogen " + "x" * 40
        first = pr.sheet_name(long, used)
        second = pr.sheet_name(long, used)
        self.assertEqual(len(first), 31)
        self.assertNotEqual(first, second)
        self.assertLessEqual(len(second), 31)

    def test_format_recipient(self):
        self.assertEqual(pr.format_recipient("Irene Gahn\nZiegenhorst 25\n23570 Lübeck"),
                         "Irene Gahn, Ziegenhorst 25, 23570 Lübeck")
        self.assertEqual(
            pr.format_recipient('[{"anrede":"Institution","name":"Amt","plz":"1","ort":"X","strassepostfach":"Str"}]'),
            "Amt, Str, 1 X")


class TestBuildSheets(PatientRecordTestBase):
    def test_missing_patients_raises(self):
        os.remove(os.path.join(self.folder, "patienten.csv"))
        with self.assertRaises(FileNotFoundError):
            pr.load_export_folder(self.folder)

    def test_missing_optional_tables_are_reported(self):
        os.remove(os.path.join(self.folder, "anatomy_patienten.csv"))
        tables, missing = pr.load_export_folder(self.folder)
        self.assertIn("anatomy_patienten", missing)
        self.assertNotIn(pr.SHEET_FINDINGS, pr.build_sheets(tables))

    def test_sheet_order(self):
        self.assertEqual(list(self.sheets()), [
            pr.SHEET_PATIENTS, pr.SHEET_EVENTS, pr.SHEET_PRESCRIPTIONS, pr.SHEET_FINDINGS,
            "Fragebogen Anamnese", pr.SHEET_NOTES, pr.SHEET_FILES,
        ])

    def test_patients(self):
        p = self.sheets()[pr.SHEET_PATIENTS]
        # Sortiert nach Nachname, Umlaute wie Grundbuchstaben (Adam < Öztürk)
        self.assertEqual(p["Nachname"].tolist(), ["Adam", "Öztürk"])
        oez = p.iloc[1]
        self.assertEqual(oez["PLZ"], "01067")  # führende Null bleibt
        self.assertEqual(oez["Zuzahlungsbefreit"], "Ja")
        self.assertTrue(pd.isna(oez["Abweichender Rechnungsempfänger"]))  # "null" -> leer
        adam = p.iloc[0]
        self.assertEqual(adam["Geburtsdatum"], "24.12.1890")
        self.assertEqual(adam["Abweichender Rechnungsempfänger"], "Herr Sohn Adam, Weg 2, 23570 Lübeck")

    def test_events(self):
        e = self.sheets()[pr.SHEET_EVENTS]
        self.assertEqual(len(e), 2)  # Flexzeit ohne Patient fällt weg
        self.assertEqual(e.columns[:4].tolist(), pr.PATIENT_REF_COLS)
        self.assertEqual(e["Datum"].tolist(), [datetime.date(2026, 1, 2), datetime.date(2026, 1, 5)])
        last = e.iloc[1]
        self.assertEqual(last["Beginn"], "10:20")
        self.assertEqual(last["Therapeut"], "Uwe Therapeut")
        self.assertEqual(last["Verordnung (Rezept-Nr)"], "152")
        self.assertEqual(last["Wahrgenommen"], "Ja")
        self.assertEqual(last["Dokumentation"], "Mobilisation HWS")
        self.assertEqual(e.iloc[0]["Nicht erschienen"], "Ja")

    def test_prescriptions(self):
        z = self.sheets()[pr.SHEET_PRESCRIPTIONS].iloc[0]
        self.assertEqual(z["Kostenträgerart"], "Gesetzlich (GKV)")
        self.assertEqual(z["Heilmittel"], "KG (6x); MT (6x)")
        self.assertEqual(z["Leitsymptomatik"], "Schädigung der Bewegungssegmente")
        self.assertEqual(z["Arzt-Nr (LANR)"], "089585001")
        self.assertEqual(z["Anzahl Termine"], 2)

    def test_findings(self):
        f = self.sheets()[pr.SHEET_FINDINGS].iloc[0]
        self.assertEqual(f["Ansicht"], "Muskulatur vorne (weiblich)")
        self.assertEqual(f["Befund"], "Fraktur")

    def test_questionnaire_one_column_per_question(self):
        q = self.sheets()["Fragebogen Anamnese"]
        # Fragen in Reihenfolge der Sortierung, Checkbox -> Ja/Nein
        self.assertEqual(q.columns.tolist()[-2:], ["Schmerzstärke", "Rauchen Sie?"])
        self.assertEqual(q.iloc[0]["Rauchen Sie?"], "Ja")
        self.assertEqual(q.iloc[0]["Schmerzstärke"], "4")
        self.assertEqual(q.iloc[0]["Nachname"], "Adam")

    def test_notes_and_files_resolve_patient(self):
        s = self.sheets()
        notes = s[pr.SHEET_NOTES]
        self.assertEqual(len(notes), 1)  # Praxis-Notiz ohne Patientenbezug fällt weg
        self.assertEqual(notes.iloc[0]["Bezug"], "Patient")
        files = s[pr.SHEET_FILES]
        self.assertEqual(len(files), 1)  # Praxis-Logo fällt weg
        self.assertEqual(files.iloc[0]["Bezug"], "Verordnung 152")
        self.assertEqual(files.iloc[0]["Nachname"], "Öztürk")
        self.assertEqual(files.iloc[0]["Kategorie"], "Rezept Vorderseite")
        self.assertEqual(files.iloc[0]["Größe (KB)"], 2)


class TestExportWorkbook(PatientRecordTestBase):
    def test_export_writes_readable_workbook(self):
        out = os.path.join(self.folder, "akte.xlsx")
        result = pr.export_patient_record(self.folder, out)
        self.assertTrue(os.path.exists(out))
        self.assertEqual(result.missing_tables, [])
        self.assertEqual(result.sheet_counts[pr.SHEET_EVENTS], 2)

        book = pd.read_excel(out, sheet_name=None, dtype=str)
        self.assertEqual(list(book)[0], pr.SHEET_OVERVIEW)
        self.assertEqual(book[pr.SHEET_PATIENTS]["PLZ"].tolist(), ["23570", "01067"])
        # Überlanger Therapiebericht wird auf Excel-Zellgröße gekürzt statt abzubrechen
        report = book[pr.SHEET_PRESCRIPTIONS]["Therapiebericht"].iloc[0]
        self.assertLessEqual(len(report), 32767)
        self.assertTrue(report.endswith("[…gekürzt]"))

    def test_text_starting_with_equals_stays_text(self):
        """Regression: Doku "=> Besserung" wurde als (ungültige) Formel geschrieben -> Excel-Reparatur."""
        import zipfile
        path = os.path.join(self.folder, "events.csv")
        with open(path, encoding="utf-8") as f:
            content = f.read()
        content = content.replace('"Mobilisation HWS"', '"=> deutliche Besserung"')
        content = content.replace('"Befund"', '"=SUMME(A1:A3)"')
        content = content.replace('"KG";"R1";"P1";"0"', '"https://example.org";"R1";"P1";"0"')
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)

        out = os.path.join(self.folder, "akte.xlsx")
        pr.export_patient_record(self.folder, out)
        with zipfile.ZipFile(out) as z:
            for name in z.namelist():
                if name.startswith("xl/worksheets/"):
                    xml = z.read(name).decode("utf-8")
                    self.assertNotIn("<f>", xml, name)
                    self.assertNotIn("<hyperlink", xml, name)
        e = pd.read_excel(out, sheet_name=pr.SHEET_EVENTS, dtype=str)
        self.assertIn("=> deutliche Besserung", e["Dokumentation"].tolist())
        self.assertIn("=SUMME(A1:A3)", e["Stichwort"].tolist())
        self.assertIn("https://example.org", e["Behandlung"].tolist())

    def test_export_cp1252_folder(self):
        """Exporte in Windows-Kodierung werden ebenfalls gelesen."""
        path = os.path.join(self.folder, "patienten.csv")
        with open(path, encoding="utf-8") as f:
            content = f.read().replace("Ayşe", "Anne")  # ş gibt es in cp1252 nicht
        with open(path, "w", encoding="cp1252") as f:
            f.write(content)
        p = self.sheets()[pr.SHEET_PATIENTS]
        self.assertIn("Öztürk", p["Nachname"].tolist())


if __name__ == "__main__":
    unittest.main()

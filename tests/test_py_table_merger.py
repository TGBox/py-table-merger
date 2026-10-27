# type: ignore
import unittest
from unittest.mock import patch, MagicMock
import tkinter as tk
import pandas as pd
import sys
import os

# Wir fügen den "src"-Ordner zum Python-Suchpfad hinzu
# Das sorgt dafür, dass 'constants' und 'table_merger' gefunden werden
current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.abspath(os.path.join(current_dir, "..", "src"))
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)

import unittest
from unittest.mock import patch, MagicMock
import tkinter as tk
import pandas as pd

# JETZT kannst du die Klassen direkt aus dem Paket importieren:
from table_merger.py_table_merger import TableMergerApp, RenameColumnsDialog

class TestTableMergerApp(unittest.TestCase):
    def setUp(self):
        """Wird vor jedem einzelnen Test ausgeführt. Initialisiert die GUI."""
        self.root = tk.Tk()
        # Mocken der constants, falls sie nicht existieren oder abweichen
        with patch('table_merger.py_table_merger.constants') as mock_const:
            mock_const.DEFAULT_WINDOW_WIDTH = 800
            mock_const.DEFAULT_WINDOW_HEIGHT = 600
            mock_const.MIN_WINDOW_WIDTH = 600
            mock_const.MIN_WINDOW_HEIGHT = 400
            mock_const.EXPORT_FG_COL = "green"
            mock_const.CONFIRM_BTN_BG_COL = "blue"
            mock_const.CONFIRM_BTN_ACTIVE_BG_COL = "lightblue"
            mock_const.CONFIRM_BTN_ACTIVE_FG_COL = "white"
            mock_const.PRIM_KEY_SELECT_BG_COL = "gold"
            
            self.app = TableMergerApp(self.root)

        # Beispiel-Datenframes für die Tests erstellen
        self.test_df_left = pd.DataFrame({
            "ID": [1, 2, 3],
            "Name": ["Alice", "Bob", "Charlie"],
            "Alter": [25, 30, 35]
        })
        self.test_df_right = pd.DataFrame({
            "KundenID": [1, 2, 4],
            "Stadt": ["Berlin", "München", "Hamburg"],
            "Alter": [25, 99, 40] # Gleicher Spaltenname wie links für Konflikttest
        })

    def tearDown(self):
        """Wird nach jedem Test ausgeführt. Schließt die GUI-Ressourcen."""
        self.root.destroy()

    def test_read_file_csv(self):
        """Testet, ob die CSV-Erkennung und das Einlesen korrekt funktionieren."""
        # Erstelle eine temporäre CSV-Datei
        temp_csv = "temp_test.csv"
        self.test_df_left.to_csv(temp_csv, sep=";", index=False)
        
        try:
            df_read = self.app.read_file(temp_csv)
            self.assertEqual(list(df_read.columns), ["ID", "Name", "Alter"])
            self.assertEqual(len(df_read), 3)
        finally:
            if os.path.exists(temp_csv):
                os.remove(temp_csv)

    @patch('tkinter.filedialog.askopenfilename')
    def test_load_left_table(self, mock_askfilename):
        """Testet das Laden der linken Tabelle und das Befüllen der Listbox."""
        mock_askfilename.return_value = "dummy_left.csv"
        
        # Wir mocken read_file, damit keine echte Datei existieren muss
        self.app.read_file = MagicMock(return_value=self.test_df_left)
        
        self.app.load_left()
        
        # Prüfen, ob die Widgets aktualisiert wurden
        self.assertEqual(self.app.lbl_left_name.cget("text"), "dummy_left.csv")
        self.assertEqual(self.app.listbox_left.size(), 3)
        self.assertEqual(self.app.listbox_left.get(0), "ID")

    def test_update_export_list_suffixes(self):
        """Testet, ob bei namensgleichen Spalten die Suffixe (Links)/(Rechts) korrekt gesetzt werden."""
        # Simulierte geladene Spalten in den Listboxen
        self.app.listbox_left.insert(tk.END, "ID", "Name", "Alter")
        self.app.listbox_right.insert(tk.END, "KundenID", "Stadt", "Alter")
        
        # Selektion simulieren ("Alter" in beiden Tabellen ausgewählt)
        self.app.listbox_left.selection_set(2)  # Alter
        self.app.listbox_right.selection_set(2) # Alter
        
        self.app.update_export_list()
        
        export_items = self.app.listbox_export.get(0, tk.END)
        self.assertIn("Alter (Links)", export_items)
        self.assertIn("Alter (Rechts)", export_items)

    @patch('tkinter.filedialog.asksaveasfilename')
    @patch('tkinter.messagebox.showinfo')
    @patch('table_merger.py_table_merger.RenameColumnsDialog')
    def test_perform_merge_success(self, mock_rename_dialog, mock_showinfo, mock_asksave):
        """Testet den gesamten Merge-Ablauf inklusive Spaltenumbenennung."""
        # Setup App-Zustand
        self.app.df_left = self.test_df_left
        self.app.df_right = self.test_df_right
        self.app.key_left = "ID"
        self.app.key_right = "KundenID"
        
        # Spaltenauswahl in Listboxen simulieren
        self.app.listbox_left.insert(tk.END, "ID", "Name")
        self.app.listbox_right.insert(tk.END, "KundenID", "Stadt")
        self.app.listbox_left.selection_set(0, 1) # ID, Name
        self.app.listbox_right.selection_set(0, 1) # KundenID, Stadt
        
        # Export-Reihenfolge simulieren
        self.app.listbox_export.insert(tk.END, "ID", "Name", "Stadt")
        
        # Mocks konfigurieren
        save_path = "test_output.csv"
        mock_asksave.return_value = save_path
        
        # Simuliere, dass der Benutzer im Umbenennen-Dialog "ID" zu "Kunden_Nummer" ändert
        mock_dialog_instance = MagicMock()
        mock_dialog_instance.result = {"ID": "Kunden_Nummer", "Name": "Name", "Stadt": "Stadt"}
        mock_rename_dialog.return_value = mock_dialog_instance
        
        try:
            self.app.perform_merge()
            
            # Prüfen, ob die Datei erstellt wurde
            self.assertTrue(os.path.exists(save_path))
            
            # Inhalt der Datei prüfen
            df_result = pd.read_csv(save_path, sep=";")
            self.assertIn("Kunden_Nummer", df_result.columns)
            self.assertNotIn("ID", df_result.columns)
            self.assertEqual(df_result.iloc[0]["Kunden_Nummer"], 1)
            self.assertEqual(df_result.iloc[0]["Stadt"], "Berlin")
            
            mock_showinfo.assert_called_once() # Erfolgsmeldung angezeigt?
            
        finally:
            if os.path.exists(save_path):
                os.remove(save_path)

    def test_read_file_keeps_text_values(self):
        """PLZ/Telefon mit führender Null bleiben Text, NULL wird leer, andere Werte bleiben erhalten."""
        temp_csv = "temp_text.csv"
        with open(temp_csv, "w", encoding="utf-8") as f:
            f.write('id;plz;tel;ik;name\n"00042";"01067";"01729239280";NULL;"NA"\n"00043";"21442";NULL;"103411401";"Müller"\n')
        try:
            df = self.app.read_file(temp_csv)
            self.assertEqual(df.loc[0, "id"], "00042")
            self.assertEqual(df.loc[0, "plz"], "01067")
            self.assertEqual(df.loc[0, "tel"], "01729239280")
            self.assertEqual(df.loc[1, "ik"], "103411401")
            self.assertTrue(pd.isna(df.loc[0, "ik"]))
            self.assertEqual(df.loc[0, "name"], "NA")
        finally:
            os.remove(temp_csv)

    def test_read_file_cp1252(self):
        """Exporte in Windows-Kodierung (cp1252) werden korrekt eingelesen."""
        temp_csv = "temp_cp1252.csv"
        with open(temp_csv, "w", encoding="cp1252") as f:
            f.write("id;name\n1;Stödter\n2;Weiß\n")
        try:
            df = self.app.read_file(temp_csv)
            self.assertEqual(df["name"].tolist(), ["Stödter", "Weiß"])
        finally:
            os.remove(temp_csv)

    def test_export_item_to_internal(self):
        """Anzeigenamen der Export-Liste werden der richtigen Tabellenseite zugeordnet."""
        f = TableMergerApp.export_item_to_internal
        self.assertEqual(f("Alter (Links)", ["Alter"]), TableMergerApp.LEFT_PREFIX + "Alter")
        self.assertEqual(f("Alter (Rechts)", ["Alter"]), TableMergerApp.RIGHT_PREFIX + "Alter")
        self.assertEqual(f("Name", ["Name"]), TableMergerApp.LEFT_PREFIX + "Name")
        self.assertEqual(f("Stadt", ["Name"]), TableMergerApp.RIGHT_PREFIX + "Stadt")

    def _run_merge(self, left, right, key_left, key_right, sel_left, sel_right, export_items):
        """Hilfsfunktion: führt perform_merge mit gemockten Dialogen aus und liefert das CSV-Ergebnis."""
        self.app.df_left, self.app.df_right = left, right
        self.app.key_left, self.app.key_right = key_left, key_right
        self.app.listbox_left.insert(tk.END, *left.columns)
        self.app.listbox_right.insert(tk.END, *right.columns)
        for c in sel_left:
            self.app.listbox_left.selection_set(list(left.columns).index(c))
        for c in sel_right:
            self.app.listbox_right.selection_set(list(right.columns).index(c))
        self.app.listbox_export.insert(tk.END, *export_items)
        save_path = "test_output_merge.csv"
        with patch("tkinter.filedialog.asksaveasfilename", return_value=save_path), \
             patch("tkinter.messagebox.showinfo"), \
             patch("table_merger.py_table_merger.RenameColumnsDialog") as dlg:
            dlg.side_effect = lambda parent, cols: MagicMock(result={c: c for c in cols})
            self.app.perform_merge()
        try:
            return pd.read_csv(save_path, sep=";", dtype=str)
        finally:
            if os.path.exists(save_path):
                os.remove(save_path)

    def test_merge_same_column_name_selected_only_left(self):
        """Regression: 'id' existiert in beiden Tabellen, wird aber nur links gewählt -> darf nicht verschwinden."""
        patienten = pd.DataFrame({"id": ["P1", "P2"], "name": ["Muster", "Test"]})
        termine = pd.DataFrame({"id": ["T1", "T2", "T3"], "patient_id": ["P1", "P1", "P2"], "start": ["a", "b", "c"]})
        df = self._run_merge(termine, patienten, "patient_id", "id",
                             sel_left=["start"], sel_right=["id", "name"],
                             export_items=["id", "name", "start"])
        self.assertEqual(list(df.columns), ["id", "name", "start"])
        self.assertEqual(df["id"].tolist(), ["P1", "P1", "P2"])

    def test_merge_key_name_collides_with_selected_column(self):
        """Regression: rechter Schlüssel heißt wie eine links gewählte Spalte, ist rechts aber nicht gewählt."""
        left = pd.DataFrame({"patient_id": ["P1", "P2"], "wert": ["x", "y"]})
        right = pd.DataFrame({"pid": ["P1", "P2"], "patient_id": ["A", "B"], "name": ["Muster", "Test"]})
        df = self._run_merge(left, right, "patient_id", "pid",
                             sel_left=["patient_id", "wert"], sel_right=["name"],
                             export_items=["patient_id", "wert", "name"])
        self.assertEqual(list(df.columns), ["patient_id", "wert", "name"])
        self.assertEqual(df["patient_id"].tolist(), ["P1", "P2"])

    def test_merge_same_column_selected_both_sides(self):
        """Gleichnamige Spalte auf beiden Seiten gewählt -> beide mit (Links)/(Rechts) im Export."""
        df = self._run_merge(self.test_df_left, self.test_df_right, "ID", "KundenID",
                             sel_left=["ID", "Alter"], sel_right=["Alter"],
                             export_items=["ID", "Alter (Links)", "Alter (Rechts)"])
        self.assertEqual(list(df.columns), ["ID", "Alter (Links)", "Alter (Rechts)"])
        self.assertEqual(df.loc[1, "Alter (Rechts)"], "99")

    @patch("tkinter.messagebox.showwarning")
    @patch("tkinter.filedialog.asksaveasfilename", return_value="test_dup.csv")
    @patch("table_merger.py_table_merger.RenameColumnsDialog")
    def test_merge_rejects_duplicate_names_after_rename(self, mock_dlg, mock_save, mock_warn):
        """Doppelte Namen im Umbenennen-Dialog führen zu einer Warnung statt zu einer kaputten Datei."""
        self.app.df_left, self.app.df_right = self.test_df_left, self.test_df_right
        self.app.key_left, self.app.key_right = "ID", "KundenID"
        self.app.listbox_left.insert(tk.END, "ID", "Name", "Alter")
        self.app.listbox_right.insert(tk.END, "KundenID", "Stadt", "Alter")
        self.app.listbox_left.selection_set(1)
        self.app.listbox_right.selection_set(1)
        self.app.listbox_export.insert(tk.END, "Name", "Stadt")
        mock_dlg.return_value = MagicMock(result={"Name": "X", "Stadt": "X"})
        self.app.perform_merge()
        mock_warn.assert_called_once()
        self.assertFalse(os.path.exists("test_dup.csv"))

    def test_excel_export_text_with_equals_sign(self):
        """Regression: Text mit führendem "=" darf im Merger-Export keine Formel werden."""
        import zipfile
        left = pd.DataFrame({"id": ["P1"], "plz": ["01067"]})
        right = pd.DataFrame({"pid": ["P1"], "doku": ["=> Besserung"]})
        self.app.df_left, self.app.df_right = left, right
        self.app.key_left, self.app.key_right = "id", "pid"
        self.app.listbox_left.insert(tk.END, "id", "plz")
        self.app.listbox_right.insert(tk.END, "pid", "doku")
        self.app.listbox_left.selection_set(1)
        self.app.listbox_right.selection_set(1)
        self.app.listbox_export.insert(tk.END, "plz", "doku")
        save_path = "test_formula.xlsx"
        with patch("tkinter.filedialog.asksaveasfilename", return_value=save_path), \
             patch("tkinter.messagebox.showinfo"), \
             patch("table_merger.py_table_merger.RenameColumnsDialog") as dlg:
            dlg.side_effect = lambda parent, cols: MagicMock(result={c: c for c in cols})
            self.app.perform_merge()
        try:
            with zipfile.ZipFile(save_path) as z:
                self.assertNotIn("<f>", z.read("xl/worksheets/sheet1.xml").decode("utf-8"))
            self.assertEqual(pd.read_excel(save_path, dtype=str).loc[0, "doku"], "=> Besserung")
        finally:
            if os.path.exists(save_path):
                os.remove(save_path)

    def test_excel_export_long_text(self):
        """Excel-Export mit langen Texten (Dokumentation) und Text-Zahlen läuft fehlerfrei durch."""
        left = pd.DataFrame({"id": ["P1"], "plz": ["01067"]})
        right = pd.DataFrame({"pid": ["P1"], "doku": ["x" * 2000]})
        self.app.df_left, self.app.df_right = left, right
        self.app.key_left, self.app.key_right = "id", "pid"
        self.app.listbox_left.insert(tk.END, "id", "plz")
        self.app.listbox_right.insert(tk.END, "pid", "doku")
        self.app.listbox_left.selection_set(1)
        self.app.listbox_right.selection_set(1)
        self.app.listbox_export.insert(tk.END, "plz", "doku")
        save_path = "test_output.xlsx"
        with patch("tkinter.filedialog.asksaveasfilename", return_value=save_path), \
             patch("tkinter.messagebox.showinfo"), patch("tkinter.messagebox.showerror") as err, \
             patch("table_merger.py_table_merger.RenameColumnsDialog") as dlg:
            dlg.side_effect = lambda parent, cols: MagicMock(result={c: c for c in cols})
            self.app.perform_merge()
        try:
            err.assert_not_called()
            df = pd.read_excel(save_path, dtype=str)
            self.assertEqual(df.loc[0, "plz"], "01067")
            self.assertEqual(len(df.loc[0, "doku"]), 2000)
        finally:
            if os.path.exists(save_path):
                os.remove(save_path)

    @patch("tkinter.messagebox.showinfo")
    @patch("tkinter.filedialog.asksaveasfilename", return_value="akte.xlsx")
    @patch("tkinter.filedialog.askdirectory")
    def test_patient_record_button_flow(self, mock_dir, mock_save, mock_info):
        """Button "Patientenakte": Ordner + Speicherort wählen -> Export wird mit beiden Pfaden aufgerufen."""
        import tempfile
        folder = tempfile.mkdtemp()
        open(os.path.join(folder, "Patienten.csv"), "w").close()  # Groß-/Kleinschreibung egal
        mock_dir.return_value = folder
        fake_result = MagicMock(sheet_counts={"Patienten": 3}, missing_tables=["files"])
        with patch("table_merger.py_table_merger.patient_record_export.export_patient_record",
                   return_value=fake_result) as mock_export:
            self.app.export_patient_record()
        mock_export.assert_called_once_with(folder, "akte.xlsx")
        message = mock_info.call_args[0][1]
        self.assertIn("Patienten: 3 Einträge", message)
        self.assertIn("files.csv", message)

    @patch("tkinter.messagebox.showwarning")
    @patch("tkinter.filedialog.asksaveasfilename")
    @patch("tkinter.filedialog.askdirectory")
    def test_patient_record_wrong_folder(self, mock_dir, mock_save, mock_warn):
        """Ordner ohne patienten.csv -> Hinweis, kein Speichern-Dialog."""
        import tempfile
        mock_dir.return_value = tempfile.mkdtemp()
        self.app.export_patient_record()
        mock_warn.assert_called_once()
        mock_save.assert_not_called()

if __name__ == "__main__":
    unittest.main()
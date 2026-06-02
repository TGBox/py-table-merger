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

if __name__ == "__main__":
    unittest.main()
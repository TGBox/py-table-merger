
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
#!Deprecated
#TODO: Hier noch checken was dann am Ende eigentlich mit den Tabellen gemacht wird. Wird eine gesamte erstellt? Wir wollen ja Datenreferenz selbst wählen können!

class Tooltip:
    """
    Eine Hilfsklasse zur Erstellung von Hover-Tooltips für Tkinter-Widgets.
    """
    def __init__(self, widget: tk.Widget, text_func): # type: ignore
        self.widget = widget
        self.text_func = text_func  # type: ignore # Funktion, die den anzuzeigenden Text liefert
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip) # type: ignore
        self.widget.bind("<Leave>", self.hide_tip) # type: ignore

    def show_tip(self, event=None): # type: ignore
        if self.tip_window or not self.text_func(): # type: ignore
            return
        # Positionierung des Tooltips nahe dem Mauszeiger/Widget
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)  # Rahmen und Titelleiste entfernen
        tw.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(
            tw, 
            text=self.text_func(),  # type: ignore
            justify=tk.LEFT,
            background="#FFFFE0",  # Klassisches Tooltip-Gelb
            relief=tk.SOLID, 
            borderwidth=1,
            font=("Arial", "9", "normal"), # type: ignore
            padx=5,
            pady=3
        )
        label.pack()

    def hide_tip(self, event=None): # type: ignore
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()

class TableMergerApp:
    """
    Eine grafische Tkinter-Anwendung, um zwei Tabellenexporte (CSV oder Excel)
    über frei wählbare Schlüsselspalten miteinander zu verknüpfen (Merge/Join).
    """
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Datenbank Export Merger")
        self.root.geometry("950x750")
        self.root.minsize(900, 700)

        # Speicher für geladene Datenframes und Dateipfade
        self.df_left = None
        self.df_right = None
        self.path_left = ""
        self.path_right = ""

        # --- OBERER BEREICH: Datei-Auswahl ---
        self.frame_files = tk.LabelFrame(root, text="1. Tabellen auswählen", font=("Arial", 10, "bold"))
        self.frame_files.pack(fill="x", padx=15, pady=10)

        # Konfiguration für das Raster (Grid) im Datei-Bereich
        self.frame_files.columnconfigure(1, weight=1)

        # Linke Tabelle wählen
        self.btn_load_left = tk.Button(
            self.frame_files, text="Tabelle Links wählen", command=self.load_left, width=20
        )
        self.btn_load_left.grid(row=0, column=0, padx=10, pady=8, sticky="ew")
        self.lbl_left_name = tk.Label(
            self.frame_files, text="Keine Datei ausgewählt", fg="gray", anchor="w"
        )
        self.lbl_left_name.grid(row=0, column=1, padx=10, pady=8, sticky="ew")

        # Rechte Tabelle wählen
        self.btn_load_right = tk.Button(
            self.frame_files, text="Tabelle Rechts wählen", command=self.load_right, width=20
        )
        self.btn_load_right.grid(row=1, column=0, padx=10, pady=8, sticky="ew")
        self.lbl_right_name = tk.Label(
            self.frame_files, text="Keine Datei ausgewählt", fg="gray", anchor="w"
        )
        self.lbl_right_name.grid(row=1, column=1, padx=10, pady=8, sticky="ew")

        # --- MITTLERER BEREICH: Spalten-Auswahl ---
        self.frame_columns = tk.LabelFrame(root, text="2. Matching-Spalten wählen", font=("Arial", 10, "bold"))
        self.frame_columns.pack(fill="both", expand=True, padx=15, pady=10)

        # Grid-Gewichtung für die Listen (damit sie sich bei Skalierung anpassen)
        self.frame_columns.columnconfigure(0, weight=1)
        self.frame_columns.columnconfigure(1, weight=1)
        self.frame_columns.rowconfigure(1, weight=1)

        # Linke Spaltenliste
        tk.Label(self.frame_columns, text="Spalten (Links):", font=("Arial", 9, "italic")).grid(
            row=0, column=0, padx=10, pady=(5, 2), sticky="w"
        )
        # exportselection=False verhindert das Deselektieren, wenn man in die andere Box klickt
        self.listbox_left = tk.Listbox(self.frame_columns, exportselection=False, selectmode=tk.SINGLE)
        self.listbox_left.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

        # Rechte Spaltenliste
        tk.Label(self.frame_columns, text="Spalten (Rechts):", font=("Arial", 9, "italic")).grid(
            row=0, column=1, padx=10, pady=(5, 2), sticky="w"
        )
        self.listbox_right = tk.Listbox(self.frame_columns, exportselection=False, selectmode=tk.SINGLE)
        self.listbox_right.grid(row=1, column=1, padx=10, pady=5, sticky="nsew")

        # --- UNTERER BEREICH: Join-Optionen & Button ---
        self.frame_options = tk.Frame(root)
        self.frame_options.pack(fill="x", padx=15, pady=5)
        
        # 1. Dropdown direkt in __init__ erstellen
        tk.Label(self.frame_options, text="Verknüpfungs-Typ (Join):").pack(side="left", padx=5)
        self.join_type = tk.StringVar(value="left")
        self.dropdown_join = tk.OptionMenu(
            self.frame_options, self.join_type, "left", "right", "inner", "outer"
        )
        self.dropdown_join.pack(side="left", padx=5)

        # 2. Label für die Info-Anzeige erstellen
        self.lbl_join_info = tk.Label(self.frame_options, text="", fg="gray", font=("Arial", 9, "italic"))
        self.lbl_join_info.pack(side="left", padx=10)

        # Automatische Aktualisierung der Info bei Auswahl-Wechsel (Trace)
        self.join_type.trace_add("write", self.update_join_info) # type: ignore

        # 3. Merge-Button direkt in __init__ erstellen
        self.btn_merge = tk.Button(
            root,
            text="Tabellen zusammenführen & speichern",
            bg="#4CAF50",
            fg="white",
            font=("Arial", 11, "bold"),
            command=self.perform_merge,
            relief="raised",
            pady=8
        )
        self.btn_merge.pack(fill="x", padx=15, pady=15)
        
        # Tooltip-Klasse für das Dropdown aktivieren
        Tooltip(self.dropdown_join, self.get_current_tooltip_text)
        
        # Text beim Start einmalig initialisieren
        self.update_join_info() # type: ignore
    

    # Beschreibungen für die jeweiligen Join-Typen
    JOIN_DESCRIPTIONS = {
        "left": "Left Join:\nBehält alle Zeilen der linken Tabelle und ergänzt passende Daten der rechten Tabelle. Fehlende Werte werden leer (NaN) gelassen.",
        "right": "Right Join:\nBehält alle Zeilen der rechten Tabelle und ergänzt passende Daten der linken Tabelle. Fehlende Werte werden leer (NaN) gelassen.",
        "inner": "Inner Join:\nHiermit können Tabellen gemerged werden, wenn bei beiden Tabellen ein passender Eintrag vorhanden ist.",
        "outer": "Outer Join:\nBehält alle Zeilen aus beiden Tabellen. Nicht übereinstimmende Zeilen werden mit leeren Werten (NaN) aufgefüllt."
    }


    def get_current_tooltip_text(self):
        """Liefert die Erklärung für den aktuell ausgewählten Join-Typ."""
        val = self.join_type.get()
        return self.JOIN_DESCRIPTIONS.get(val, "")

    def update_join_info(self, *args):  # type: ignore
        """Aktualisiert NUR die Textanzeige neben dem Dropdown-Menü."""
        desc = self.get_current_tooltip_text()
        compact_desc = desc.replace("\n", " -> ")
        self.lbl_join_info.config(text=compact_desc)

    def read_file(self, filepath: str):
        """
        Versucht eine CSV- oder Excel-Datei intelligent in ein Pandas DataFrame einzulesen.
        Probiert bei CSVs typische Trennzeichen aus.
        """
        ext = os.path.splitext(filepath)[1].lower()
        if ext == ".csv":
            # Testet verschiedene Standard-Trennzeichen für CSV-Dateien
            for sep in [";", ",", "\t"]:
                try:
                    df = pd.read_csv(filepath, sep=sep)
                    # Wenn nur eine Spalte eingelesen wird, war das Trennzeichen wahrscheinlich falsch
                    if len(df.columns) > 1:
                        return df
                except Exception:
                    continue
            # Letzter Versuch mit Standard-Read
            return pd.read_csv(filepath)
        elif ext in [".xlsx", ".xls"]:
            excel_file: pd.DataFrame = pd.read_excel(filepath) # type: ignore
            return excel_file
        else:
            raise ValueError("Ungültiges Dateiformat. Bitte CSV- oder Excel-Dateien verwenden.")

    def load_left(self):
        """Öffnet den File-Explorer für die linke Tabelle und befüllt die linke Liste."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Tabellendaten", "*.csv *.xlsx *.xls")]
        )
        if file_path:
            try:
                self.df_left = self.read_file(file_path)
                self.path_left = file_path
                self.lbl_left_name.config(
                    text=os.path.basename(file_path), fg="black"
                )

                # Linke Listbox mit den neuen Spaltennamen füllen
                self.listbox_left.delete(0, tk.END)
                for col in self.df_left.columns:
                    self.listbox_left.insert(tk.END, col)
            except Exception as e:
                messagebox.showerror("Ladefehler", f"Linke Datei konnte nicht gelesen werden:\n{e}")

    def load_right(self):
        """Öffnet den File-Explorer für die rechte Tabelle und befüllt die rechte Liste."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Tabellendaten", "*.csv *.xlsx *.xls")]
        )
        if file_path:
            try:
                self.df_right = self.read_file(file_path)
                self.path_right = file_path
                self.lbl_right_name.config(
                    text=os.path.basename(file_path), fg="black"
                )

                # Rechte Listbox mit den neuen Spaltennamen füllen
                self.listbox_right.delete(0, tk.END)
                for col in self.df_right.columns:
                    self.listbox_right.insert(tk.END, col)
            except Exception as e:
                messagebox.showerror("Ladefehler", f"Rechte Datei konnte nicht gelesen werden:\n{e}")

    def perform_merge(self):
        """Führt die Verknüpfung der Tabellen basierend auf der Auswahl durch."""
        # Prüfen, ob beide Dateien geladen wurden
        if self.df_left is None or self.df_right is None:
            messagebox.showwarning(
                "Dateien fehlen", "Bitte wähle zuerst beide Tabellendateien aus!"
            )
            return

        # Prüfen, ob in beiden Listen eine Spalte ausgewählt ist
        sel_left = self.listbox_left.curselection() # type: ignore
        sel_right = self.listbox_right.curselection() # type: ignore

        if not sel_left or not sel_right:
            messagebox.showwarning(
                "Spaltenauswahl fehlt", "Bitte markiere in beiden Spaltenlisten den jeweiligen Schlüssel!"
            )
            return

        # Spaltennamen aus den Selektionen ziehen
        col_left = self.listbox_left.get(sel_left[0]) # type: ignore
        col_right = self.listbox_right.get(sel_right[0]) # type: ignore

        # Dialog zum Speichern der Datei öffnen
        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV-Datei (Semikolon-getrennt)", ".csv"), ("Excel-Arbeitsmappe", ".xlsx")],
            title="Zusammengeführte Tabelle speichern unter..."
        )

        if save_path:
            try:
                # Merge mit dem ausgewählten Join-Typ (left, right, inner, outer) durchführen
                df_result = pd.merge(
                    self.df_left,
                    self.df_right,
                    left_on=col_left, # type: ignore
                    right_on=col_right, # type: ignore
                    how=self.join_type.get(), # type: ignore
                )

                # Speichern im gewünschten Format (basierend auf Dateiendung)
                if save_path.endswith(".xlsx"):
                    df_result.to_excel(save_path, index=False) # type: ignore
                else:
                    # Speichern mit Semikolon (Standard für deutsche Excel-Versionen)
                    df_result.to_csv(save_path, index=False, sep=";")

                messagebox.showinfo(
                    "Erfolg", f"Zusammenführung erfolgreich!\n\nDatei gespeichert unter:\n{save_path}"
                )
            except Exception as e:
                messagebox.showerror(
                    "Merge-Fehler", f"Ein Fehler ist während des Zusammenführens aufgetreten:\n{e}"
                )


# Hauptprogramm starten
if __name__ == "__main__":
    root: tk.Tk = tk.Tk()
    app: TableMergerApp = TableMergerApp(root)
    root.mainloop()
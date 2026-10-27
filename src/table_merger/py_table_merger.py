#type: ignore
import os
import sys

# Berechne den Pfad zum "src"-Ordner (eine Ebene über dieser Datei)
current_dir = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.abspath(os.path.join(current_dir, ".."))

# Füge "src" zu den Suchpfaden hinzu
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)

import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import src.constants as constants
import src.table_io as table_io
from src.patient_record import patient_record_export

class RenameColumnsDialog(tk.Toplevel):
    def __init__(self, parent, columns):
        super().__init__(parent)
        self.title("Spalten umbenennen")
        
        # Feste Startgröße definieren
        self.geometry("500x450")
        self.minsize(400, 300)
        
        # Modal machen (blockiert das Hauptfenster)
        self.transient(parent)
        self.grab_set()
        
        self.result = None
        self.entries = {}
        
        # --- HAUPTLAYOUT ---
        # Wir nutzen pack() für die Hauptelemente, um die Buttons fest unten zu verankern.
        
        # 1. Überschrift / Info-Label ganz oben
        lbl_info = tk.Label(
            self, 
            text="Passen Sie hier die Spaltennamen für den Export an:", 
            font=("Roboto", 10, "italic"), 
            fg="gray30"
        )
        lbl_info.pack(fill="x", padx=15, pady=(10, 5), anchor="w")
        
        # 2. Button-Frame ganz unten FIXIEREN (wichtig: zuerst packen oder side="bottom" nutzen)
        btn_frame = tk.Frame(self)
        btn_frame.pack(fill="x", side="bottom", padx=15, pady=15)
        
        btn_cancel = tk.Button(btn_frame, text="Abbrechen", command=self.on_cancel, width=10)
        btn_cancel.pack(side="right", padx=5)
        
        btn_ok = tk.Button(
            btn_frame,
            text="Übernehmen",
            command=self.on_ok,
            width=12,
            bg=constants.CONFIRM_BTN_BG_COL,
            activebackground=constants.CONFIRM_BTN_ACTIVE_BG_COL,
            fg=constants.CONFIRM_BTN_FG_COL,
            activeforeground=constants.CONFIRM_BTN_ACTIVE_FG_COL,
            font=("Roboto", 9, "bold")
        )
        btn_ok.pack(side="right", padx=5)
        
        # 3. Der mittlere, scrollbare Bereich (erhält den restlichen Platz durch expand=True)
        container = tk.Frame(self)
        container.pack(fill="both", expand=True, padx=15, pady=5)
        
        canvas = tk.Canvas(container, borderwidth=0, highlightthickness=0)
        scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
        
        self.scrollable_frame = tk.Frame(canvas)
        
        # Grid-Gewichtung im scrollable_frame, damit Spalten gleichmäßig verteilt werden
        self.scrollable_frame.columnconfigure(0, weight=1)
        self.scrollable_frame.columnconfigure(1, weight=1)
        
        # Event: Wenn sich die Größe des inneren Frames ändert, Scrollregion anpassen
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # WICHTIG: Das Fenster im Canvas zentrieren und an die Canvas-Breite koppeln
        canvas_window = canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        # Sorgt dafür, dass das innere Panel genauso breit wird wie der Canvas selbst
        canvas.bind('<Configure>', lambda event: canvas.itemconfig(canvas_window, width=event.width))
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # --- SPALTEN UND EINGABEFELDER IM GRID ---
        # Spaltenüberschriften im scrollbaren Bereich
        tk.Label(self.scrollable_frame, text="Ursprünglicher Name", font=("Roboto", 10, "bold")).grid(row=0, column=0, padx=10, pady=8, sticky="w")
        tk.Label(self.scrollable_frame, text="Neuer Name", font=("Roboto", 10, "bold")).grid(row=0, column=1, padx=10, pady=8, sticky="ew")
        
        for idx, col in enumerate(columns, start=1):
            lbl = tk.Label(self.scrollable_frame, text=col, fg="gray20", font=("Roboto", 9, "bold"), anchor="w")
            lbl.grid(row=idx, column=0, padx=10, pady=6, sticky="w")
            
            entry = tk.Entry(self.scrollable_frame, font=("Roboto", 10))
            entry.insert(0, col)  # Standardmäßig alten Namen einfügen
            entry.grid(row=idx, column=1, padx=10, pady=6, sticky="ew")
            
            self.entries[col] = entry
            
        # Mausrad-Scrolling aktivieren
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))

        self.wait_window(self)

    def on_ok(self):
        # Mapping von altem Namen zu neuem Namen erstellen
        self.result = {}
        for old_name, entry in self.entries.items():
            new_name = entry.get().strip()
            # Falls das Feld leer gelassen wurde, behalten wir den alten Namen
            self.result[old_name] = new_name if new_name else old_name
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()

class Tooltip:
    def __init__(self, widget: tk.Widget, text_func):
        self.widget = widget
        self.text_func = text_func
        self.tip_window = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text_func():
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(
            tw, text=self.text_func(), justify=tk.LEFT, background="#F2FFE0",
            relief=tk.SOLID, borderwidth=1, font=("Roboto", "9", "normal"), padx=5, pady=3
        )
        label.pack()

    def hide_tip(self, event=None):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()

class DragDropListbox(tk.Listbox):
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.bind('<Button-1>', self.on_start_drag)
        self.bind('<B1-Motion>', self.on_drag)
        self.cur_index = None

    def on_start_drag(self, event):
        self.cur_index = self.nearest(event.y)

    def on_drag(self, event):
        i = self.nearest(event.y)
        if i != self.cur_index and self.cur_index is not None:
            text = self.get(self.cur_index)
            # Zustand (z.B. Selektion) merken, falls nötig, hier reicht Text-Tausch
            self.delete(self.cur_index)
            self.insert(i, text)
            self.selection_clear(0, tk.END)
            self.selection_set(i)
            self.cur_index = i

class TableMergerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Datenbank Export Merger")
        # Calculating the offset to center the window in the primary screen.
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (constants.DEFAULT_WINDOW_WIDTH // 2)
        y = (screen_height // 2) - (constants.DEFAULT_WINDOW_HEIGHT // 2)
        self.root.geometry(f"{constants.DEFAULT_WINDOW_WIDTH}x{constants.DEFAULT_WINDOW_HEIGHT}+{x}+{y}")
        self.root.minsize(constants.MIN_WINDOW_WIDTH, constants.MIN_WINDOW_HEIGHT)
        

        self.df_left = None
        self.df_right = None
        
        # Speicher für die per Rechtsklick gewählten Primärschlüssel
        self.key_left = None
        self.key_right = None

        # --- 1. DATEI-AUSWAHL ---
        self.frame_files = tk.LabelFrame(root, text="1. Tabellen auswählen", font=("Roboto", 10, "bold"))
        self.frame_files.pack(fill="x", padx=15, pady=10)
        self.frame_files.columnconfigure(1, weight=1)

        self.btn_load_left = tk.Button(
            self.frame_files,
            text="Tabelle Links wählen",
            command=self.load_left,
            width=20,
            bg="#5E5E5E",
            fg="#cccccc",
            activebackground="#444444",
            activeforeground="#ffffff",
            font=("Roboto", 11, "bold")
        )
        self.btn_load_left.grid(row=0, column=0, padx=10, pady=8, sticky="ew")
        self.lbl_left_name = tk.Label(self.frame_files, text="Keine Datei ausgewählt", fg="gray", anchor="w")
        self.lbl_left_name.grid(row=0, column=1, padx=10, pady=8, sticky="ew")

        self.btn_load_right = tk.Button(
            self.frame_files,
            text="Tabelle Rechts wählen",
            command=self.load_right,
            width=20,
            bg="#5E5E5E",
            fg="#cccccc",
            activebackground="#444444",
            activeforeground="#ffffff",
            font=("Roboto", 11, "bold")
        )
        self.btn_load_right.grid(row=0, column=3, padx=10, pady=8, sticky="ew")
        self.lbl_right_name = tk.Label(self.frame_files, text="Keine Datei ausgewählt", fg="gray", anchor="w")
        self.lbl_right_name.grid(row=0, column=2, padx=10, pady=8, sticky="ew")

        # --- 2. SPALTEN-AUSWAHL ---
        self.frame_columns = tk.LabelFrame(root, text="2. Spalten wählen (Rechtsklick = Schlüssel | Linksklick = Export-Spalten)", font=("Roboto", 10, "bold"))
        self.frame_columns.pack(fill="both", expand=True, padx=15, pady=10)

        self.frame_columns.columnconfigure(0, weight=3)
        self.frame_columns.columnconfigure(1, weight=3)
        self.frame_columns.columnconfigure(2, weight=1)
        self.frame_columns.rowconfigure(1, weight=1)

        # Linke Tabelle
        tk.Label(self.frame_columns, text="Spalten (Links):", font=("Roboto", 9, "italic")).grid(row=0, column=0, padx=10, pady=(5, 2), sticky="w")
        self.listbox_left = tk.Listbox(self.frame_columns, exportselection=False, selectmode=tk.MULTIPLE)
        self.listbox_left.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")
        self.listbox_left.bind("<Button-3>", self.set_key_left) # Rechtsklick Event
        self.listbox_left.bind("<<ListboxSelect>>", self.update_export_list) # Event bei Auswahl-Änderung

        # Rechte Tabelle
        tk.Label(self.frame_columns, text="Spalten (Rechts):", font=("Roboto", 9, "italic")).grid(row=0, column=1, padx=10, pady=(5, 2), sticky="w")
        self.listbox_right = tk.Listbox(self.frame_columns, exportselection=False, selectmode=tk.MULTIPLE)
        self.listbox_right.grid(row=1, column=1, padx=10, pady=5, sticky="nsew")
        self.listbox_right.bind("<Button-3>", self.set_key_right) # Rechtsklick Event
        self.listbox_right.bind("<<ListboxSelect>>", self.update_export_list) # Event bei Auswahl-Änderung
        
        # Drag and Drop Sortierbereich
        tk.Label(self.frame_columns, text="Finale Export-Reihenfolge (Ziehen zum Sortieren):", font=("Roboto", 9, "bold"), fg=constants.EXPORT_FG_COL).grid(row=0, column=2, padx=10, pady=(5, 2), sticky="w")
        self.listbox_export = DragDropListbox(self.frame_columns, background="#F0FBF0", selectbackground="#B8D6D8")
        self.listbox_export.grid(row=1, column=2, padx=10, pady=5, sticky="nsew")

        # --- 3. OPTIONSBEREICH ---
        self.frame_options = tk.Frame(root)
        self.frame_options.pack(fill="x", padx=15, pady=5)
        
        tk.Label(self.frame_options, text="Verknüpfungs-Typ (Join):").pack(side="left", padx=5)
        self.join_type = tk.StringVar(value="inner")
        self.dropdown_join = tk.OptionMenu(self.frame_options, self.join_type, "left", "right", "inner", "outer")
        self.dropdown_join.pack(side="left", padx=5)

        self.lbl_join_info = tk.Label(self.frame_options, text="", fg="gray", font=("Roboto", 9, "italic"))
        self.lbl_join_info.pack(side="left", padx=10)
        self.join_type.trace_add("write", self.update_join_info)

        self.btn_merge = tk.Button(
            root, text="Ausgewählte Spalten zusammenführen & speichern",
            bg=constants.CONFIRM_BTN_BG_COL, fg="white", font=("Roboto", 11, "bold"),
            activebackground=constants.CONFIRM_BTN_ACTIVE_BG_COL, activeforeground=constants.CONFIRM_BTN_ACTIVE_FG_COL,
            command=self.perform_merge, relief="raised", pady=8
        )
        self.btn_merge.pack(fill="x", padx=15, pady=(15, 5))

        # --- 4. PATIENTENAKTE (fester Export aus komplettem CSV-Export) ---
        self.btn_patient_record = tk.Button(
            root, text="Patientenakte aus komplettem Export erstellen …",
            bg="#5E5E5E", fg="#F7F7F7", font=("Roboto", 10, "bold"),
            activebackground="#444444", activeforeground="#ffffff",
            command=self.export_patient_record, relief="raised", pady=6
        )
        self.btn_patient_record.pack(fill="x", padx=15, pady=(5, 15))
        Tooltip(self.btn_patient_record, lambda: (
            "Erstellt aus dem Ordner mit ALLEN CSV-Tabellen einer Instanz eine Excel-Arbeitsmappe:\n"
            "Patienten, Termine & Dokumentation, Verordnungen, Befunde, Fragebögen, Notizen, Dokumente."
        ))
        
        Tooltip(self.dropdown_join, self.get_current_tooltip_text)
        self.update_join_info()

    JOIN_DESCRIPTIONS = {
        "left": "Left Join:\nBehält alle Zeilen links, ergänzt passende Daten rechts.",
        "right": "Right Join:\nBehält alle Zeilen rechts, ergänzt passende Daten links.",
        "inner": "Inner Join:\nNur Zeilen, die in beiden Tabellen einen Treffer haben.",
        "outer": "Outer Join:\nBehält alle Zeilen aus beiden Tabellen."
    }

    def get_current_tooltip_text(self):
        return self.JOIN_DESCRIPTIONS.get(self.join_type.get(), "")

    def update_join_info(self, *args):
        desc = self.get_current_tooltip_text()
        self.lbl_join_info.config(text=desc.replace("\n", " -> "))
        
    def update_export_list(self, event=None):
        """Aktualisiert die 3. Spalte basierend auf den Klicks in Links und Rechts, 
        behält aber bereits bestehende Sortierungen bei."""
        # Aktuell ausgewählte Spalten ermitteln
        sel_left = [self.listbox_left.get(i) for i in self.listbox_left.curselection()]
        sel_right = [self.listbox_right.get(i) for i in self.listbox_right.curselection()]
        
        # Suffixe simulieren, die Pandas bei Namensgleichheit vergibt (_x, _y)
        # Um Konflikte zu vermeiden, benennen wir sie hier für die Anzeige eindeutig
        current_selected = []
        for c in sel_left:
            if c in sel_right:
                current_selected.append(f"{c} (Links)")
            else:
                current_selected.append(c)
                
        for c in sel_right:
            if c in sel_left:
                current_selected.append(f"{c} (Rechts)")
            else:
                current_selected.append(c)

        # Bestehende Einträge in der Export-Liste merken
        existing_items = list(self.listbox_export.get(0, tk.END))

        # 1. Entferne Einträge, die nicht mehr ausgewählt sind
        for item in existing_items[::-1]: 
            if item not in current_selected:
                idx = existing_items.index(item)
                self.listbox_export.delete(idx)
                existing_items.remove(item)

        # 2. Füge neue Einträge am Ende hinzu
        for item in current_selected:
            if item not in existing_items:
                self.listbox_export.insert(tk.END, item)

    # --- RECHTSKLICK FUNKTIONEN FÜR SCHLÜSSEL ---
    def set_key_left(self, event):
        index = self.listbox_left.nearest(event.y)
        if index >= 0:
            # Vorherige Schlüssel-Färbung zurücksetzen
            for i in range(self.listbox_left.size()):
                if self.listbox_left.itemcget(i, "bg") == constants.PRIM_KEY_SELECT_BG_COL:
                    self.listbox_left.itemconfig(i, bg="white", fg="black")
            
            self.key_left = self.listbox_left.get(index)
            self.listbox_left.itemconfig(index, bg=constants.PRIM_KEY_SELECT_BG_COL, fg="black") # Gold markieren

    def set_key_right(self, event):
        index = self.listbox_right.nearest(event.y)
        if index >= 0:
            for i in range(self.listbox_right.size()):
                if self.listbox_right.itemcget(i, "bg") == constants.PRIM_KEY_SELECT_BG_COL:
                    self.listbox_right.itemconfig(i, bg="white", fg="black")
            
            self.key_right = self.listbox_right.get(index)
            self.listbox_right.itemconfig(index, bg=constants.PRIM_KEY_SELECT_BG_COL, fg="black")

    # --- DATEN EINLESEN ---
    def read_file(self, filepath: str):
        """Liest CSV/Excel ein (alle Werte als Text, Encoding-Erkennung) – siehe table_io.read_table."""
        return table_io.read_table(filepath)

    def load_left(self):
        file_path = filedialog.askopenfilename(filetypes=[("Tabellendaten", "*.csv *.xlsx *.xls")])
        if file_path:
            try:
                self.df_left = self.read_file(file_path)
                self.lbl_left_name.config(text=os.path.basename(file_path), fg="black")
                self.listbox_left.delete(0, tk.END)
                self.key_left = None
                for col in self.df_left.columns:
                    self.listbox_left.insert(tk.END, col)
                self.btn_load_left.config(bg="#51995B", fg="#F7F7F7")
            except Exception as e:
                messagebox.showerror("Ladefehler", str(e))

    def load_right(self):
        file_path = filedialog.askopenfilename(filetypes=[("Tabellendaten", "*.csv *.xlsx *.xls")])
        if file_path:
            try:
                self.df_right = self.read_file(file_path)
                self.lbl_right_name.config(text=os.path.basename(file_path), fg="black")
                self.listbox_right.delete(0, tk.END)
                self.key_right = None
                for col in self.df_right.columns:
                    self.listbox_right.insert(tk.END, col)
                self.btn_load_right.config(bg="#51995B", fg="#F7F7F7")
            except Exception as e:
                messagebox.showerror("Ladefehler", str(e))

    # --- MERGE PROCESS ---
    # Interne Präfixe, damit beim Merge nie Namenskonflikte (und damit _x/_y-Suffixe) entstehen.
    LEFT_PREFIX = "__L__"
    RIGHT_PREFIX = "__R__"

    @classmethod
    def export_item_to_internal(cls, item: str, selected_left) -> str:
        """Übersetzt einen Eintrag der Export-Liste in den internen (präfixierten) Spaltennamen."""
        if item.endswith(" (Links)"):
            return cls.LEFT_PREFIX + item[: -len(" (Links)")]
        if item.endswith(" (Rechts)"):
            return cls.RIGHT_PREFIX + item[: -len(" (Rechts)")]
        return (cls.LEFT_PREFIX if item in selected_left else cls.RIGHT_PREFIX) + item

    def perform_merge(self):
        if self.df_left is None or self.df_right is None:
            messagebox.showwarning("Fehler", "Bitte beide Tabellen laden!")
            return

        if not self.key_left or not self.key_right:
            messagebox.showwarning("Schlüssel fehlt", "Bitte in beiden Listen einen Primärschlüssel mit RECHTSKLICK wählen!")
            return

        # Indizes für den Merge holen (unabhängig von der Export-Anzeige)
        sel_left_idx = self.listbox_left.curselection()
        sel_right_idx = self.listbox_right.curselection()

        selected_left = [self.listbox_left.get(i) for i in sel_left_idx]
        cols_left = list(selected_left)
        cols_right = [self.listbox_right.get(i) for i in sel_right_idx]

        # Die Schlüssel MÜSSEN für den Merge-Prozess temporär im DF enthalten sein
        if self.key_left not in cols_left:
            cols_left.append(self.key_left)
        if self.key_right not in cols_right:
            cols_right.append(self.key_right)

        # Alle Spalten mit Seiten-Präfix versehen -> keine Namenskonflikte beim Merge möglich.
        df_l_filtered = self.df_left[cols_left].add_prefix(self.LEFT_PREFIX)
        df_r_filtered = self.df_right[cols_right].add_prefix(self.RIGHT_PREFIX)
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV-Datei", ".csv"), ("Excel-Arbeitsmappe", ".xlsx")],
            title="Speichern unter..."
        )

        if save_path:
            try:
                # 1. Tabellen regulär über die Schlüssel zusammenführen
                df_result = pd.merge(
                    df_l_filtered,
                    df_r_filtered,
                    left_on=self.LEFT_PREFIX + self.key_left,
                    right_on=self.RIGHT_PREFIX + self.key_right,
                    how=self.join_type.get()
                )

                # 2. Gewünschte finale Reihenfolge aus der Drag&Drop-Listbox auslesen
                #    und interne Namen -> Anzeigenamen (z. B. "Alter (Links)") zuordnen.
                column_map = {}
                for item in self.listbox_export.get(0, tk.END):
                    internal = self.export_item_to_internal(item, selected_left)
                    if internal in df_result.columns:
                        column_map[internal] = item
                final_columns_ordered = list(column_map)

                # Falls der Benutzer gar keine Spalten gewählt hat, brechen wir ab
                if not final_columns_ordered:
                    messagebox.showwarning("Keine Spalten", "Es wurden keine Spalten für den Export ausgewählt.")
                    return

                # DataFrame auf die sortierte Reihenfolge einschränken
                df_result = df_result[final_columns_ordered].rename(columns=column_map)

                # --- NEU: Dialog zum Umbenennen der Spalten aufrufen ---
                rename_dialog = RenameColumnsDialog(self.root, df_result.columns)
                
                # Wenn der Benutzer den Dialog abbricht (result ist None), stoppen wir den Prozess
                if rename_dialog.result is None:
                    return
                
                # DataFrame mit den neuen Namen umbenennen
                df_result = df_result.rename(columns=rename_dialog.result)

                # Doppelte Spaltennamen nach dem Umbenennen würden den Export zerschießen
                duplicates = df_result.columns[df_result.columns.duplicated()].unique().tolist()
                if duplicates:
                    messagebox.showwarning(
                        "Doppelte Spaltennamen",
                        "Folgende Spaltennamen kommen mehrfach vor:\n" + ", ".join(map(str, duplicates))
                        + "\n\nBitte eindeutige Namen vergeben."
                    )
                    return
                # -------------------------------------------------------

                # 3. Speichern
                if save_path.endswith(".xlsx"):
                    # Formatierter Export (fette Überschriften, Spaltenbreiten, Filter) – siehe table_io
                    with table_io.excel_writer(save_path) as writer:
                        table_io.write_excel_sheet(writer, df_result, "Merge_Ergebnis")

                else:
                    # CSV unterstützt prinzipiell keine Formatierungen wie Spaltenbreiten oder "Fett"
                    df_result.to_csv(save_path, index=False, sep=";")
                    
                messagebox.showinfo("Speichern erfolgreich", f"Die neue Tabelle wurde erfolgreich unter \"{save_path}\" abgespeichert.")
            except Exception as e:
                messagebox.showerror("Merge-Fehler", str(e))

    # --- PATIENTENAKTE ---
    def export_patient_record(self):
        folder = filedialog.askdirectory(title="Ordner mit dem kompletten CSV-Export wählen")
        if not folder:
            return
        if patient_record_export.find_table_file(folder, patient_record_export.REQUIRED_TABLE) is None:
            messagebox.showwarning(
                "Falscher Ordner",
                "Im gewählten Ordner wurde keine Datei \"patienten.csv\" gefunden.\n\n"
                "Bitte den Ordner wählen, der alle CSV-Tabellen des Exports enthält."
            )
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel-Arbeitsmappe", ".xlsx")],
            initialfile=f"Patientenakte_{datetime.date.today():%Y-%m-%d}.xlsx",
            title="Patientenakte speichern unter..."
        )
        if not save_path:
            return

        self.root.config(cursor="watch")
        self.root.update_idletasks()
        try:
            result = patient_record_export.export_patient_record(folder, save_path)
        except PermissionError:
            messagebox.showerror(
                "Speichern nicht möglich",
                f"Die Datei \"{save_path}\" konnte nicht geschrieben werden.\n"
                "Ist sie vielleicht noch in Excel geöffnet?"
            )
            return
        except Exception as e:
            messagebox.showerror("Export-Fehler", f"Die Patientenakte konnte nicht erstellt werden:\n{e}")
            return
        finally:
            self.root.config(cursor="")

        lines = [f"• {name}: {count} Einträge" for name, count in result.sheet_counts.items()]
        if result.missing_tables:
            lines.append("\nNicht im Export enthalten (Blätter entfallen ggf.):\n"
                         + ", ".join(f"{t}.csv" for t in result.missing_tables))
        messagebox.showinfo(
            "Patientenakte erstellt",
            f"Gespeichert unter:\n{save_path}\n\n" + "\n".join(lines)
            + "\n\nHinweis: Die Datei enthält Gesundheitsdaten – nur verschlüsselt weitergeben."
        )


if __name__ == "__main__":
    root = tk.Tk()
    app = TableMergerApp(root)
    root.mainloop()
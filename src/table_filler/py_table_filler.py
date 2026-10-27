import csv
import tkinter as tk
from tkinter import filedialog, messagebox


def select_file():
    """Öffnet den Dialog zum Auswählen der Eingabedatei."""
    filepath = filedialog.askopenfilename(
        title="CSV-Datei auswählen",
        filetypes=[("CSV-Dateien", "*.csv"), ("Alle Dateien", "*.*")],
    )
    if filepath:
        entry_input.delete(0, tk.END)
        entry_input.insert(0, filepath)

        # Automatischen Vorschlag für den Ausgabepfad generieren
        if not entry_output.get():
            suggested_output = filepath.replace(".csv", "_mit_null.csv")
            entry_output.insert(0, suggested_output)


def select_save_path():
    """Öffnet den Dialog zum Wählen des Speicherorts."""
    filepath = filedialog.asksaveasfilename(
        title="Speicherort wählen",
        defaultextension=".csv",
        filetypes=[("CSV-Dateien", "*.csv"), ("Alle Dateien", "*.*")],
    )
    if filepath:
        entry_output.delete(0, tk.END)
        entry_output.insert(0, filepath)


def process_csv():
    """Liest die CSV ein, ersetzt leere Felder mit 'NULL' und speichert sie neu."""
    input_path = entry_input.get().strip()
    output_path = entry_output.get().strip()

    if not input_path or not output_path:
        messagebox.showwarning(
            "Fehlende Eingabe",
            "Bitte wählen Sie sowohl eine Eingabedatei als auch einen Speicherort aus.",
        )
        return

    # Liste von Codierungen, die nacheinander probiert werden
    encodings_to_try = ["utf-8", "cp1252", "iso-8859-1"]
    content = None
    used_encoding = None

    # 1. Datei mit richtiger Kodierung einlesen
    for encoding in encodings_to_try:
        try:
            with open(input_path, mode="r", encoding=encoding) as infile:
                content = infile.read()
                used_encoding = encoding
                break
        except UnicodeDecodeError:
            continue

    if content is None:
        messagebox.showerror(
            "Fehler",
            "Die Datei konnte mit keiner gängigen Zeichenkodierung gelesen werden.",
        )
        return

    try:
        # 2. Trennzeichen (Delimiter) automatisch erkennen
        try:
            # Testet die ersten 2048 Zeichen auf Trennzeichen wie ; oder ,
            dialect = csv.Sniffer().sniff(content[:2048])
        except csv.Error:
            # Falls die Erkennung fehlschlägt, nutzen wir das Semikolon als Standard
            dialect = csv.excel
            dialect.delimiter = ";"

        lines = content.splitlines()

        # 3. Schreiben mit demselben Trennzeichen wie das Original
        with open(
            output_path, mode="w", encoding="utf-8", newline=""
        ) as outfile:
            reader = csv.reader(lines, dialect=dialect)
            writer = csv.writer(outfile, dialect=dialect)

            zeilen_zaehler = 0
            null_zaehler = 0

            for row in reader:
                neue_zeile = []
                for val in row:
                    # Prüft ob das Feld leer ist (oder nur aus Leerzeichen/Anführungszeichen besteht)
                    clean_val = val.strip().replace('"', "")
                    if clean_val == "":
                        neue_zeile.append("NULL")
                        null_zaehler += 1
                    else:
                        neue_zeile.append(val)

                writer.writerow(neue_zeile)
                zeilen_zaehler += 1

        messagebox.showinfo(
            "Erfolg!",
            f"Datei erfolgreich verarbeitet!\n\n"
            f"Erkanntes Trennzeichen: '{dialect.delimiter}'\n"
            f"Ersetzte NULL-Werte: {null_zaehler}\n"
            f"Verarbeitete Zeilen: {zeilen_zaehler}\n\n"
            f"Gespeichert unter:\n{output_path}",
        )

    except Exception as e:
        messagebox.showerror(
            "Fehler", f"Es ist ein Fehler aufgetreten:\n{str(e)}"
        )


# --- GUI Erstellen ---
root = tk.Tk()
root.title("CSV NULL-Füller")
root.geometry("550x220")
root.resizable(False, False)

# Main Container mit Padding
frame = tk.Frame(root, padx=15, pady=15)
frame.pack(fill=tk.BOTH, expand=True)

# 1. Eingabedatei
lbl_input = tk.Label(frame, text="Quelldatei (CSV):", anchor="w")
lbl_input.pack(fill=tk.X, pady=(0, 2))

frame_input = tk.Frame(frame)
frame_input.pack(fill=tk.X, pady=(0, 10))

entry_input = tk.Entry(frame_input)
entry_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

btn_browse_input = tk.Button(
    frame_input, text="Durchsuchen...", command=select_file
)
btn_browse_input.pack(side=tk.RIGHT)

# 2. Ausgabedatei
lbl_output = tk.Label(frame, text="Speicherort für neue Datei:", anchor="w")
lbl_output.pack(fill=tk.X, pady=(0, 2))

frame_output = tk.Frame(frame)
frame_output.pack(fill=tk.X, pady=(0, 15))

entry_output = tk.Entry(frame_output)
entry_output.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

btn_browse_output = tk.Button(
    frame_output, text="Durchsuchen...", command=select_save_path
)
btn_browse_output.pack(side=tk.RIGHT)

# 3. Start-Button
btn_process = tk.Button(
    frame,
    text="Leere Felder mit NULL befüllen",
    bg="#28a745",
    fg="white",
    font=("Arial", 10, "bold"),
    pady=8,
    command=process_csv,
)
btn_process.pack(fill=tk.X)

# GUI starten
root.mainloop()
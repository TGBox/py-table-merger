# Datenbank Export Merger 📊✨

Ein benutzerfreundliches Python-Tool mit grafischer Oberfläche (Tkinter), um zwei Tabellen (CSV- oder Excel-Daten) intelligent über Primärschlüssel miteinander zu verknüpfen, Spalten flexibel zu filtern, per Drag-and-Drop zu sortieren und vor dem Export umzubenennen.

---

## 🎯 Einsatzzweck

Beim Arbeiten mit Datenbank-Exporten steht man oft vor dem Problem, dass Daten über mehrere Berichte oder Tabellen verteilt sind (z. B. Kundendaten in Tabelle A, Bestellungen in Tabelle B). 

Dieses Tool wurde entwickelt, um **ohne SQL- oder Programmierkenntnisse** und **ohne aufgeblähte Tabellenkalkulationsprogramme**:
* Zwei Datenexporte schnell und präzise über Schlüsselspalten zu verknüpfen (Join).
* Nur die Spalten auszuwählen, die für den finalen Bericht wirklich relevant sind.
* Die finale Reihenfolge der Spalten intuitiv festzulegen.
* Kryptische Datenbank-Spaltennamen vor dem Speichern in "schöne", lesbare Bezeichner umzubenennen.

---

## 🚀 Funktionen

* **Format-Flexibilität:** Unterstützt das Einlesen und Schreiben von `.csv`, `.xlsx` und `.xls` Dateien.
* **Intelligenter CSV-Import:** Erkennt Trennzeichen (`;`, `,`, `\t`) automatisch.
* **Visuelle Schlüsselwahl:** Primärschlüssel werden einfach per **Rechtsklick** in den Spaltenlisten definiert.
* **Flexible Join-Typen:** Unterstützt *Left Join*, *Right Join*, *Inner Join* und *Outer Join* inklusive Live-Erklärung in der Oberfläche.
* **Drag-and-Drop Sortierung:** Die finale Export-Reihenfolge der Spalten kann per Maus verschoben werden.
* **Interaktiver Umbenennen-Dialog:** Vor dem Speichern öffnet sich ein scrollbares Fenster, in dem alle gewählten Spalten manuell umbenannt werden können.
* **Erweiterter Excel-Export:** Beim Speichern als `.xlsx` werden Spaltenbreiten automatisch an den Inhalt angepasst und Überschriften fett formatiert.

---

## 🛠️ Projektstruktur

Das Projekt ist modular nach modernen Python-Standards aufgebaut:

```text
py_table_merger/
├── src/
│   ├── constants.py            # Zentrale Design- und Konfigurationswerte
│   └── table_merger/
│       ├── __init__.py
│       └── py_table_merger.py  # Hauptanwendung (GUI & Logik)
├── tests/
│   └── test_py_table_merger.py # Automatisierte Testsuite
└── README.md                   # Projektdokumentation
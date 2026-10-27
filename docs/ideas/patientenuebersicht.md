# Patientenübersicht (Export ohne IDs)

## Problem Statement

Wie können wir einer Praxis ihre Patienten, deren Aktivität, mögliche Dubletten und den Terminverlauf mit Dokumentation in einer Excel-Datei geben, die ganz ohne technische IDs lesbar und navigierbar ist?

## Recommended Direction

Eigener Button **„Patientenübersicht (ohne IDs) erstellen …“** im Merger. Im Code teilen sich Patientenakte und Patientenübersicht **eine Pipeline mit zwei Profilen**; das Profil bestimmt die Blätter und ob Patienten über Datenbank-IDs oder eine **laufende Nr.** referenziert werden. Die Lfd. Nr. wird nach deutscher Namenssortierung vergeben und verbindet alle Blätter.

Die Lfd. Nr. ist ein **interner Link** (Patientenliste → erster Termin; Termine und Dubletten → Patientenzeile). Ja/Nein-Werte, Datumsformate, Spaltenbreiten und Filter wie in der Patientenakte.

Verworfen wurde ein separat gebauter Export (Doppelpflege des Dubletten-Blatts; Fixes wie der „=“-Formel-Fehler müssten an zwei Stellen gemacht werden).

## Key Assumptions to Validate

- [ ] Die Praxis braucht keine exportübergreifende Referenz – *beim Kunden nachfragen, ob mit der Datei Tickets bearbeitet werden.*
- [ ] Links, die nach dem Umsortieren falsch springen, sind mit Hinweis akzeptabel – *Testdatei sortieren, Hinweis auf Verständlichkeit prüfen.*
- [ ] Name, Geburtsdatum, Anschrift und Aktivitätszahlen reichen, um Dubletten im PVS zu finden – *Tahkol/Tehakol im Testsystem damit auffinden.*

## MVP Scope

- **Übersicht:** Erstellungsdatum, Einträge je Blatt, Datenschutz- und Link-Hinweis
- **Patienten:** Lfd. Nr. (Link), Anrede, Nachname, Vorname, Geburtsdatum, Straße, Hausnr., PLZ, Ort, Telefon, Telefon geschäftlich, Mobil, E-Mail, Krankenkasse, Versichertennr., Privat, Zuzahlungsbefreit, Angelegt am, Letzter Termin, Anzahl Termine, Anzahl Verordnungen, Anzahl Rechnungen (ohne Stornos)
- **Mögliche Dubletten:** Logik der Patientenakte, Lfd. Nr. (Link) statt Patienten-Nr.
- **Termine & Dokumentation:** Lfd. Nr. (Link zurück), Name, Geburtsdatum, Datum, Beginn, Ende, Dauer, Behandlung, Therapeut, Rezept-Nr., Termin der Serie, Wahrgenommen, Nicht erschienen, Abgesagt, Hausbesuch, Abgerechnet, Stichwort, Dokumentation, Bemerkung, Memo
- Patientenakte auf die gemeinsame Pipeline umgestellt – Ergebnis unverändert (zellgenau verglichen)
- Tests für Profile, Zählungen, Links und Abwesenheit von IDs

## Not Doing (and Why)

- **Gruppierte/einklappbare Terminliste** – zerstört Filtern und Sortieren in Excel
- **Ein Blatt pro Patient** – ab einigen hundert Patienten unbrauchbar
- **Datenauskunft für einen einzelnen Patienten (Art. 15 DSGVO)** – eigener Anwendungsfall, später über Patientenauswahl
- **Rechnungen nach Status aufgeteilt** – Entscheidung: eine Zahl ohne Stornos
- **Stabile Nummern über mehrere Exporte** – bräuchte wieder eine ID
- **Verordnungen, Befunde, Fragebögen in dieser Datei** – dafür gibt es die Patientenakte

## Open Questions

- ~~Rücksprung-Link im Termine-Blatt?~~ → umgesetzt
- ~~Dateiname?~~ → `Patientenuebersicht_<Datum>.xlsx`

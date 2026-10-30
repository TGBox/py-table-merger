"""Lemniscus-Export: überführt den kompletten CSV-Export einer Kundeninstanz in das
Importformat der Kundenvorlage "akten-vorlage.xlsx" (Blätter Akten, Notizen, Termine)
plus ein zusätzliches Blatt "Versicherung" (Krankenkasse/Versichertennummer).

Verknüpfung der Blätter über "Nummer" bzw. "Akte Nummer" (laufende Nr., keine Datenbank-IDs).
Die Logik ist GUI-frei; Helfer (Einlesen, Datums-/Ja-Nein-Umwandlung, Nummerierung)
kommen aus dem Patientenakte-Export.
"""
import datetime
import json
import os
import re
from dataclasses import dataclass, field

import pandas as pd
import xlsxwriter

import src.table_io as table_io
from src.patient_record import patient_record_export as pr

# --------------------------------------------------------------------------- #
# Vorlage (Spalten, Breiten und Formate exakt wie in akten-vorlage.xlsx)
# --------------------------------------------------------------------------- #
SHEET_FILES = "Akten"
SHEET_NOTES = "Notizen"
SHEET_EVENTS = "Termine"
SHEET_INSURANCE = "Versicherung"  # Zusatzblatt (nicht Teil der Vorlage)

FILE_COLUMNS = [
    "Nummer", "Vorname", "Nachname", "Akten-Typ", "Titel", "Geburtsdatum", "Geschlecht", "Beruf",
    "Firmenbezeichnung", "Firmenkontakt", "Diagnose", "Straße", "PLZ", "Stadt", "Land",
    "Telefon Privat", "Telefon Büro", "Telefon Mobil", "Fax", "Email", "Anmerkungen", "LKF",
    "Aktiv", "Status", "Rechnung Anrede", "Rechnung Vorname", "Rechnung Nachname",
    "Rechnung Titel", "Rechnung Firma", "Rechnung Zusatz", "Rechnung Straße", "Rechnung PLZ",
    "Rechnung Stadt", "Rechnung Land", "Rechnung Geburtsdatum", "Rechnung Email",
]
NOTE_COLUMNS = ["Akte Nummer", "Titel", "Zeitpunkt", "Text"]
EVENT_COLUMNS = [
    "Akte Nummer", "Terminbezeichnung", "Beginn", "Dauer", "Diagnose", "Raum",
    "Wahrgenommen", "Fertig bearbeitet", "Anmerkungen",
]
INSURANCE_COLUMNS = [
    "Akte Nummer", "Nachname", "Vorname", "Geburtsdatum", "Krankenkasse", "IK Krankenkasse",
    "Versichertennummer", "Versichertenstatus", "Privatversichert", "Zuzahlungsbefreit",
    "Zuzahlungsbefreit bis",
]

# Spaltenbreiten aus der Vorlage (Excel-XML-Werte; fehlende = Excel-Standard)
FILE_WIDTHS = {
    "A": 8.83, "D": 11.66, "E": 14.33, "F": 14.33, "I": 19.5, "J": 14.5, "K": 19.66, "L": 19.66,
    "O": 8.83, "P": 13.5, "Q": 13.16, "R": 14.0, "U": 16.66, "V": 16.66, "W": 11.5, "X": 10.33,
    "Y": 19.0, "Z": 18.83, "AA": 21.0, "AB": 16.0, "AC": 16.66, "AD": 17.33, "AE": 17.33,
    "AF": 15.5, "AG": 15.5, "AH": 15.33, "AI": 24.16, "AJ": 15.83,
}
NOTE_WIDTHS = {"A": 15.33, "C": 19.0, "D": 66.5}
EVENT_WIDTHS = {"A": 15.5, "B": 19.33, "C": 19.0, "E": 18.33, "F": 7.33, "G": 16.0, "H": 16.33, "I": 29.33}
INSURANCE_WIDTHS = {"A": 15.33, "B": 18.0, "C": 18.0, "D": 14.33, "E": 30.0, "F": 16.0, "G": 20.0,
                    "H": 18.0, "I": 16.0, "J": 18.0, "K": 22.0}
# xlsxwriter schlägt auf die angegebene Breite ein Zeichen-Polster (5 px / 7 px) auf
_WIDTH_PADDING = 5 / 7

# Spalten mit besonderem Zahlenformat
DATE_COLUMNS = {"Geburtsdatum", "Rechnung Geburtsdatum", "Zuzahlungsbefreit bis"}  # dd.mm.yyyy
DATETIME_COLUMNS = {"Zeitpunkt", "Beginn"}                                          # dd.mm.yyyy hh:mm
TEXT_FORMAT_COLUMNS = {(SHEET_NOTES, "Text")}                                       # Zellformat "@"
# In der Vorlage blau hinterlegte Kopfzellen – vermutlich die Pflichtfelder des Imports
HEADER_FILL = "#88BEF9"
REQUIRED_COLUMNS = {
    SHEET_FILES: ["Nummer", "Vorname", "Nachname"],
    SHEET_NOTES: ["Titel"],
    SHEET_EVENTS: ["Terminbezeichnung", "Beginn", "Dauer"],
}
FILLED_HEADERS = {(sheet, col) for sheet, cols in REQUIRED_COLUMNS.items() for col in cols}

# Feste Werte
FILE_TYPE = "Patient"
COUNTRY = "DE"
ACTIVE_STATUS = "Aktiv"
DEFAULT_ROOM = 1
DEFAULT_EVENT_TITLE = "Termin"
GENDER_BY_SALUTATION = {"herr": "m", "frau": "w"}

NOTE_TITLE_EVENT_DOC = "Termindokumentation"
NOTE_TITLE_FINDING = "Befund"
NOTE_TITLE_REPORT = "Therapiebericht"


@dataclass(frozen=True)
class _Profile:
    title: str
    file_prefix: str


# Anzeigename + Dateiname-Vorschlag für die Oberfläche
PROFILE = _Profile(title="Lemniscus Export", file_prefix="Lemniscus_Export")


@dataclass
class LemniscusResult:
    out_path: str
    sheet_counts: dict = field(default_factory=dict)
    missing_tables: list = field(default_factory=list)
    skipped_events: int = 0            # Termine ohne (bekannten) Patienten
    missing_required: dict = field(default_factory=dict)  # Blatt -> {Pflichtspalte: Anzahl leerer Zeilen}
    unparsed_recipients: list = field(default_factory=list)  # Akte-Nummern mit unklarem Rechnungsempfänger

    @property
    def hints(self) -> list:
        """Hinweise für die Erfolgsmeldung in der Oberfläche."""
        hints = []
        for sheet, cols in self.missing_required.items():
            for col, count in cols.items():
                hints.append(f"{sheet}: {count} Zeile(n) ohne Pflichtfeld \u201e{col}\u201c \u2013 "
                             "werden beim Import ggf. abgelehnt.")
        if self.skipped_events:
            hints.append(f"{self.skipped_events} Termine ohne bekannten Patienten wurden übersprungen.")
        if self.unparsed_recipients:
            hints.append("Rechnungsempfänger bitte prüfen (Adresse nicht erkannt) bei Akte Nr.: "
                         + ", ".join(map(str, self.unparsed_recipients)))
        return hints


# --------------------------------------------------------------------------- #
# Wert-Umwandlungen
# --------------------------------------------------------------------------- #
def _val(v):
    """Leere Werte (NaN/None/"") einheitlich zu None."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, str) and not v.strip():
        return None
    return v


def normalize_phone(value):
    """Deutsche Telefonnummer -> internationales Format wie in der Vorlage ("+49123456789").
    "0451 27798" -> "+4945127798", "0049 172 1" -> "+491721", "+49 (0)172 1" -> "+491721".
    Nicht erkennbare Angaben (z. B. mit Text) bleiben unverändert."""
    value = _val(value)
    if value is None:
        return None
    text = str(value).strip()
    compact = re.sub(r"[\s/\-().]", "", text.replace("(0)", ""))
    if not re.fullmatch(r"\+?\d{5,}", compact):
        return text
    if compact.startswith("+"):
        return compact
    if compact.startswith("00"):
        return "+" + compact[2:]
    if compact.startswith("0"):
        return "+49" + compact[1:]
    return compact  # ohne Vorwahl – nicht raten


def to_datetime_value(value, with_time: bool):
    """ISO-Text -> datetime (für Excel). Ungültig/vor 1900 -> None."""
    value = _val(value)
    if value is None:
        return None
    ts = pd.to_datetime(value, errors="coerce", format="mixed")
    if pd.isna(ts) or ts.year < 1900:
        return None
    ts = ts.to_pydatetime().replace(second=0, microsecond=0)
    return ts if with_time else ts.replace(hour=0, minute=0)


def to_bool(value) -> bool:
    """1/t/true/ja -> True, alles andere (inkl. leer) -> False."""
    value = _val(value)
    return value is not None and str(value).strip().lower() in pr._YES


def street_with_number(street, number):
    street, number = _val(street), _val(number)
    if street is None:
        return None if number is None else str(number)
    if number is None or str(street).rstrip().endswith(str(number)):
        return str(street)
    return f"{street} {number}"


def diagnosis_text(rezept_row) -> str:
    """Diagnosetext der Verordnung, ersatzweise ICD-10-Code(s)."""
    text = _val(rezept_row.get("diagnose"))
    if text:
        return str(text)
    codes = [str(c) for c in (_val(rezept_row.get("icd10_1")), _val(rezept_row.get("icd10_2"))) if c]
    return ", ".join(codes) or None


def parse_recipient(value) -> dict:
    """Abweichender Rechnungsempfänger (Freitext oder JSON) -> Rechnung-Spalten der Vorlage.

    JSON  [{"anrede","name","strassepostfach","plz","ort"}]: Personen-Name "Nachname Vorname"
          (Schreibweise des PVS), anrede "Institution" -> Rechnung Firma.
    Text  "Vorname Nachname\\nStraße Nr\\nPLZ Ort[\\nZusatz…]".
    Gibt {} zurück, wenn nichts erkennbar ist; "_unklar": True markiert Best-Effort-Zuordnungen."""
    value = _val(value)
    if value is None:
        return {}
    text = str(value).strip()
    entry = None
    if text.startswith("[") or text.startswith("{"):
        try:
            data = json.loads(text)
            entry = data[0] if isinstance(data, list) and data else data
        except ValueError:
            entry = None
    out = {}
    if isinstance(entry, dict):
        name = _val(entry.get("name")) or _val(entry.get("vorname"))
        salutation = _val(entry.get("anrede"))
        if salutation in ("Herr", "Frau"):
            out["Rechnung Anrede"] = salutation
            parts = str(name or "").split()
            if len(parts) >= 2:
                out["Rechnung Nachname"], out["Rechnung Vorname"] = parts[0], " ".join(parts[1:])
            elif parts:
                out["Rechnung Nachname"] = parts[0]
        elif name:
            out["Rechnung Firma"] = str(name)
        out["Rechnung Straße"] = _val(entry.get("strassepostfach"))
        out["Rechnung PLZ"] = _val(entry.get("plz"))
        out["Rechnung Stadt"] = _val(entry.get("ort"))
    else:
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        if not lines:
            return {}
        city_idx = next((i for i, l in enumerate(lines) if re.match(r"^\d{4,5}\s+\S", l)), None)
        name = lines[0]
        parts = name.split()
        if len(parts) >= 2:
            out["Rechnung Vorname"], out["Rechnung Nachname"] = " ".join(parts[:-1]), parts[-1]
        else:
            out["Rechnung Nachname"] = name
        if city_idx is not None:
            plz, city = lines[city_idx].split(None, 1)
            out["Rechnung PLZ"], out["Rechnung Stadt"] = plz, city
            street_lines = lines[1:city_idx]
            if street_lines:
                out["Rechnung Straße"] = street_lines[-1]
                if len(street_lines) > 1:
                    out["Rechnung Zusatz"] = ", ".join(street_lines[:-1])
            extra = lines[city_idx + 1:]
            if extra:
                out["Rechnung Zusatz"] = ", ".join(filter(None, [out.get("Rechnung Zusatz")] + extra))
        else:
            out["_unklar"] = True
            out["Rechnung Zusatz"] = ", ".join(lines[1:]) or None
    out = {k: v for k, v in out.items() if _val(v) is not None}
    if any(k in out for k in ("Rechnung Straße", "Rechnung PLZ", "Rechnung Stadt")):
        out["Rechnung Land"] = COUNTRY
    return out


# --------------------------------------------------------------------------- #
# Blätter
# --------------------------------------------------------------------------- #
def file_numbers(tables) -> dict:
    """Patienten-ID -> laufende Nummer (1..n, alphabetisch – wie in der Patientenübersicht)."""
    ref = pr.patient_reference(tables["patienten"], anonymous=True)
    return dict(zip(ref["__pid"], ref[pr.RUNNING_NO_COL].astype(int)))


def _latest_diagnosis_by_patient(tables) -> dict:
    if "rezepte" not in tables:
        return {}
    z = tables["rezepte"].copy()
    z["__sort"] = pr._col(z, "datum").fillna("")
    z = z.sort_values("__sort", kind="stable")
    result = {}
    for _, row in z.iterrows():
        text = diagnosis_text(row)
        if text:
            result[row["patienten_id"]] = text  # spätere Verordnungen überschreiben frühere
    return result


def build_files(tables, numbers) -> tuple:
    """Blatt "Akten". Gibt (DataFrame, [Nummern mit unklarem Rechnungsempfänger]) zurück."""
    p = tables["patienten"].drop_duplicates("id")
    diagnoses = _latest_diagnosis_by_patient(tables)
    rows, unclear = [], []
    for _, r in p.iterrows():
        nr = numbers[r["id"]]
        recipient = parse_recipient(r.get("rechnungsempfaenger"))
        if recipient.pop("_unklar", False):
            unclear.append(nr)
        salutation = str(_val(r.get("p_anrede")) or "").strip().lower()
        row = {
            "Nummer": nr,
            "Vorname": _val(r.get("p_vname")),
            "Nachname": _val(r.get("p_name")),
            "Akten-Typ": FILE_TYPE,
            "Geburtsdatum": to_datetime_value(r.get("p_birth"), with_time=False),
            "Geschlecht": GENDER_BY_SALUTATION.get(salutation),
            "Firmenkontakt": False,
            "Diagnose": diagnoses.get(r["id"]),
            "Straße": street_with_number(r.get("p_street"), r.get("p_hausnummer")),
            "PLZ": _val(r.get("p_plz")),
            "Stadt": _val(r.get("p_ort")),
            "Land": COUNTRY if any(_val(r.get(c)) for c in ("p_street", "p_plz", "p_ort")) else None,
            "Telefon Privat": normalize_phone(r.get("p_tel")),
            "Telefon Büro": normalize_phone(r.get("p_telge")),
            "Telefon Mobil": normalize_phone(r.get("p_handy")),
            "Email": _val(r.get("p_email")),
            "Aktiv": True,
            "Status": ACTIVE_STATUS,
            **recipient,
        }
        rows.append(row)
    df = pd.DataFrame(rows, columns=FILE_COLUMNS).sort_values("Nummer", kind="stable")
    return df.reset_index(drop=True), sorted(unclear)


def _event_prescriptions(tables) -> dict:
    """Termin-ID -> Verordnungs-Zeile (erste zugeordnete Verordnung)."""
    if "rezepte_events" not in tables or "rezepte" not in tables:
        return {}
    rx = {row["id"]: row for _, row in tables["rezepte"].iterrows()}
    re_ = tables["rezepte_events"].dropna(subset=["events_id"])
    result = {}
    for ev, rid in zip(re_["events_id"], re_["rezepte_id"]):
        if ev not in result and rid in rx:
            result[ev] = rx[rid]
    return result


def build_events(tables, numbers) -> tuple:
    """Blatt "Termine". Gibt (DataFrame, Anzahl übersprungener Termine ohne bekannten Patienten) zurück."""
    if "events" not in tables:
        return pd.DataFrame(columns=EVENT_COLUMNS), 0
    e = tables["events"]
    therapists = {}
    if "resources" in tables:
        therapists = dict(zip(tables["resources"]["id"], tables["resources"]["name"]))
    prescriptions = _event_prescriptions(tables)

    rows, skipped = [], 0
    for _, r in e.iterrows():
        pid = _val(r.get("patient_id"))
        if pid is None:
            continue  # Blocker wie "Flexzeit" – keine Akte
        if pid not in numbers:
            skipped += 1
            continue
        rx = prescriptions.get(r["id"])
        notes = []
        if to_bool(r.get("abgesagt")):
            notes.append("Abgesagt")
        if to_bool(r.get("nicht_wahrgenommen")):
            notes.append("Nicht erschienen")
        if to_bool(r.get("hausbesuch")):
            notes.append("Hausbesuch")
        therapist = therapists.get(_val(r.get("resource")))
        if therapist:
            notes.append(f"Therapeut: {therapist}")
        if rx is not None and _val(rx.get("rezeptnr")):
            series = _val(r.get("eventsequenz"))
            notes.append(f"Verordnung {rx['rezeptnr']}" + (f", Termin {series}" if series else ""))
        for col in ("description", "memo"):
            if _val(r.get(col)):
                notes.append(str(r[col]).strip())
        start = to_datetime_value(r.get("start"), with_time=True)
        duration = pd.to_numeric(_val(r.get("duration")), errors="coerce")
        if pd.isna(duration):  # Pflichtfeld: ersatzweise aus Ende - Beginn
            end = to_datetime_value(r.get("end"), with_time=True)
            if start and end and end > start:
                duration = (end - start).total_seconds() // 60
        rows.append({
            "Akte Nummer": numbers[pid],
            "Terminbezeichnung": _val(r.get("title")) or DEFAULT_EVENT_TITLE,
            "Beginn": start,
            "Dauer": int(duration) if pd.notna(duration) else None,
            "Diagnose": diagnosis_text(rx) if rx is not None else None,
            "Raum": DEFAULT_ROOM,
            "Wahrgenommen": to_bool(r.get("wahrgenommen")),
            "Fertig bearbeitet": to_bool(r.get("abgerechnet")),
            "Anmerkungen": "\n".join(notes) or None,
            "__sort": str(_val(r.get("start")) or ""),
        })
    df = pd.DataFrame(rows, columns=EVENT_COLUMNS + ["__sort"])
    df = df.sort_values(["Akte Nummer", "__sort"], kind="stable").drop(columns="__sort")
    return df.reset_index(drop=True), skipped


def build_notes(tables, numbers) -> pd.DataFrame:
    """Blatt "Notizen": Termindokumentation, Notizen, Befunde, Therapieberichte, Fragebögen."""
    notes = []

    def add(pid, title, when, text):
        text = _val(text)
        if pid in numbers and text is not None:
            notes.append({"Akte Nummer": numbers[pid], "Titel": title,
                          "Zeitpunkt": to_datetime_value(when, with_time=True), "Text": str(text),
                          "__sort": str(_val(when) or "")})

    # 1. Termindokumentation (Zeitpunkt = Terminbeginn, wie im Beispiel)
    if "events" in tables:
        for _, r in tables["events"].iterrows():
            keyword, doc = _val(r.get("dok_stichwort")), _val(r.get("dokumentation"))
            if keyword is None and doc is None:
                continue
            text = "\n\n".join(x for x in [f"Stichwort: {keyword}" if keyword else None, doc] if x)
            add(_val(r.get("patient_id")), NOTE_TITLE_EVENT_DOC, r.get("start"), text)

    # 2. Notizen (history) – auch solche an Verordnungen/Terminen, über den Patientenbezug
    if "history" in tables:
        resolver = pr._RefResolver(tables)
        h = tables["history"]
        pids, refs = resolver.apply(h)
        for (_, r), pid, ref in zip(h.iterrows(), pids, refs):
            title = _val(r.get("stichwort")) or "Notiz"
            parts = [_val(r.get("text"))]
            if ref and ref != "Patient":
                parts.append(f"(Bezug: {ref})")
            if _val(r.get("bearbeiter")):
                parts.append(f"Bearbeiter: {r['bearbeiter']}")
            add(pid, str(title), r.get("datum"), "\n".join(str(p) for p in parts if p))

    # 3. Befunde aus dem Körperschema
    if "anatomy_patienten" in tables:
        views = {}
        if "anatomy" in tables:
            views = dict(zip(tables["anatomy"]["id"], tables["anatomy"]["name"].map(pr.view_label)))
        for _, r in tables["anatomy_patienten"].iterrows():
            title = _val(r.get("title"))
            view = views.get(_val(r.get("id_region")))
            text = "\n".join(str(x) for x in [title, f"Ansicht: {view}" if view else None,
                                              _val(r.get("notes"))] if x)
            add(_val(r.get("id_patient")), f"{NOTE_TITLE_FINDING}: {title}" if title else NOTE_TITLE_FINDING,
                r.get("timestamp"), text)

    # 4. Therapieberichte aus Verordnungen
    if "rezepte" in tables:
        for _, r in tables["rezepte"].iterrows():
            report = _val(r.get("tp_therapiebericht"))
            if report is None:
                continue
            nr = _val(r.get("rezeptnr"))
            outcome = _val(r.get("tp_erfolg"))
            text = report + (f"\n\nTherapieerfolg: {outcome}" if outcome else "")
            add(r["patienten_id"], f"{NOTE_TITLE_REPORT} Verordnung {nr}" if nr else NOTE_TITLE_REPORT,
                _val(r.get("tp_erstellt_am")) or r.get("datum"), text)

    # 5. Fragebögen: "Frage: Antwort" je Zeile (direkt über die Patienten-ID zugeordnet)
    needed = ("kartei_history", "kartei_content", "pat_property")
    if all(n in tables for n in needed):
        props = tables["pat_property"].copy()
        props["__sort"] = pd.to_numeric(pr._col(props, "sortierung"), errors="coerce").fillna(0)
        props = props.sort_values("__sort", kind="stable")
        labels = dict(zip(props["id"], props["label"].fillna("Frage")))
        order = {pid: i for i, pid in enumerate(props["id"])}
        types = dict(zip(props["id"], pr._col(props, "proptyp")))
        names = {}
        if "kartei" in tables:
            names = dict(zip(tables["kartei"]["id"], tables["kartei"]["karteiname"]))
        content = tables["kartei_content"]
        for _, h in tables["kartei_history"].iterrows():
            answers = content[content["kartei_history_id"] == h["id"]]
            answers = answers.assign(__o=answers["kartei_property_id"].map(order)).sort_values("__o")
            lines = []
            for prop, val in zip(answers["kartei_property_id"], answers["content"]):
                if _val(val) is None or prop not in labels:
                    continue
                if types.get(prop) == "CHK":
                    val = "Ja" if to_bool(val) else "Nein"
                lines.append(f"{labels[prop]}: {val}")
            kartei = _val(names.get(h.get("kartei_id")))
            add(_val(h.get("patienten_id")), f"Fragebogen {kartei}" if kartei else "Fragebogen",
                h.get("datum"), "\n".join(lines))

    df = pd.DataFrame(notes, columns=NOTE_COLUMNS + ["__sort"])
    df = df.sort_values(["Akte Nummer", "__sort"], kind="stable").drop(columns="__sort")
    return df.reset_index(drop=True)


def build_insurance(tables, numbers) -> pd.DataFrame:
    """Zusatzblatt "Versicherung" (in der Vorlage nicht vorgesehen)."""
    p = tables["patienten"].drop_duplicates("id")
    df = pd.DataFrame({
        "Akte Nummer": p["id"].map(numbers),
        "Nachname": pr._col(p, "p_name"),
        "Vorname": pr._col(p, "p_vname"),
        "Geburtsdatum": pr._col(p, "p_birth").map(lambda v: to_datetime_value(v, with_time=False)),
        "Krankenkasse": pr._col(p, "p_krankenkasse"),
        "IK Krankenkasse": pr._col(p, "p_ik"),
        "Versichertennummer": pr._col(p, "p_vnr"),
        "Versichertenstatus": pr._col(p, "p_vs"),
        "Privatversichert": pr._col(p, "p_privatversichert").map(to_bool),
        "Zuzahlungsbefreit": pr._col(p, "p_zuzahlungsbefreit").map(to_bool),
        "Zuzahlungsbefreit bis": pr._col(p, "p_zuzahlungsbefreit_bis").map(
            lambda v: to_datetime_value(v, with_time=False)),
    })
    return df.sort_values("Akte Nummer", kind="stable").reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Schreiben im Vorlagen-Format
# --------------------------------------------------------------------------- #
def _write_sheet(workbook, name, df, widths):
    ws = workbook.add_worksheet(name)
    base = {"font_name": "Arial", "font_size": 10}
    header = workbook.add_format({**base, "bold": True})
    header_filled = workbook.add_format({**base, "bold": True, "bg_color": HEADER_FILL})
    plain = workbook.add_format(base)
    date_fmt = workbook.add_format({**base, "num_format": "dd.mm.yyyy"})
    datetime_fmt = workbook.add_format({**base, "num_format": "dd.mm.yyyy hh:mm"})
    text_fmt = workbook.add_format({**base, "num_format": "@"})
    limit = table_io.MAX_EXCEL_CELL_LEN - len(table_io.TRUNCATION_NOTE)

    for col_letter, width in widths.items():
        idx = _col_index(col_letter)
        ws.set_column(idx, idx, width - _WIDTH_PADDING)
    for c, title in enumerate(df.columns):
        ws.write_string(0, c, title, header_filled if (name, title) in FILLED_HEADERS else header)

    for r, row in enumerate(df.itertuples(index=False), start=1):
        for c, value in enumerate(row):
            value = _val(value)
            if value is None:
                continue
            if hasattr(value, "item") and not isinstance(value, (str, bytes)):
                value = value.item()  # numpy-Skalare -> Python-Typen
            column = df.columns[c]
            if isinstance(value, bool):
                ws.write_boolean(r, c, value, plain)
            elif isinstance(value, (datetime.datetime, datetime.date)):
                fmt = datetime_fmt if column in DATETIME_COLUMNS else date_fmt
                if not isinstance(value, datetime.datetime):
                    value = datetime.datetime(value.year, value.month, value.day)
                ws.write_datetime(r, c, value, fmt)
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                ws.write_number(r, c, value, plain)
            else:
                text = str(value)
                if len(text) > table_io.MAX_EXCEL_CELL_LEN:
                    text = text[:limit] + table_io.TRUNCATION_NOTE
                ws.write_string(r, c, text, text_fmt if (name, column) in TEXT_FORMAT_COLUMNS else plain)
    return ws


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def build_lemniscus(tables) -> tuple:
    """Alle Blätter (Name -> DataFrame) + Zusatzinfos für das Ergebnis."""
    numbers = file_numbers(tables)
    files, unclear = build_files(tables, numbers)
    events, skipped = build_events(tables, numbers)
    sheets = {
        SHEET_FILES: files,
        SHEET_NOTES: build_notes(tables, numbers),
        SHEET_EVENTS: events,
        SHEET_INSURANCE: build_insurance(tables, numbers),
    }
    return sheets, skipped, unclear


def write_lemniscus(sheets: dict, out_path: str) -> dict:
    widths = {SHEET_FILES: FILE_WIDTHS, SHEET_NOTES: NOTE_WIDTHS,
              SHEET_EVENTS: EVENT_WIDTHS, SHEET_INSURANCE: INSURANCE_WIDTHS}
    # Text bleibt immer Text (keine Formeln/Links aus Inhalten wie "=> Besserung")
    workbook = xlsxwriter.Workbook(out_path, {"strings_to_formulas": False, "strings_to_urls": False})
    try:
        for name, df in sheets.items():
            _write_sheet(workbook, name, df, widths[name])
    finally:
        workbook.close()
    return {name: len(df) for name, df in sheets.items()}


def export_lemniscus(folder: str, out_path: str) -> LemniscusResult:
    """Komplettablauf: Export-Ordner lesen -> Lemniscus-Format bauen -> Excel schreiben."""
    tables, missing = pr.load_export_folder(folder)
    needed = {"events", "resources", "rezepte", "rezepte_events", "history",
              "anatomy", "anatomy_patienten", "kartei", "kartei_history", "kartei_content", "pat_property"}
    sheets, skipped, unclear = build_lemniscus(tables)
    counts = write_lemniscus(sheets, out_path)
    return LemniscusResult(out_path=out_path, sheet_counts=counts,
                           missing_tables=[m for m in missing if m in needed],
                           skipped_events=skipped, unparsed_recipients=unclear,
                           missing_required=check_required(sheets))


def check_required(sheets: dict) -> dict:
    """Zählt leere Pflichtfelder je Blatt (nur Spalten mit mindestens einer Lücke)."""
    result = {}
    for sheet, cols in REQUIRED_COLUMNS.items():
        df = sheets.get(sheet)
        if df is None:
            continue
        gaps = {c: int(df[c].map(lambda v: _val(v) is None).sum()) for c in cols}
        gaps = {c: n for c, n in gaps.items() if n}
        if gaps:
            result[sheet] = gaps
    return result

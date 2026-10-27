"""Export "Patientenakte": baut aus dem kompletten CSV-Export einer Kundeninstanz
eine lesbare Excel-Arbeitsmappe mit einem Blatt je Thema.

Die Logik ist bewusst GUI-frei (reine pandas-Funktionen), damit sie testbar bleibt.
Fehlende optionale Tabellen führen nicht zum Abbruch, das jeweilige Blatt entfällt.
"""
import datetime
import difflib
import json
import os
import re
import unicodedata
from dataclasses import dataclass, field

import pandas as pd

import src.table_io as table_io

# Tabellen, die aus dem Export-Ordner gelesen werden ("patienten" ist Pflicht).
REQUIRED_TABLE = "patienten"
OPTIONAL_TABLES = [
    "events", "resources", "rezepte", "rezepte_events",
    "anatomy", "anatomy_patienten",
    "kartei", "kartei_history", "kartei_content", "pat_property",
    "history", "files",
]

# Feste Blattnamen
SHEET_OVERVIEW = "Übersicht"
SHEET_PATIENTS = "Patienten"
SHEET_DUPLICATES = "Mögliche Dubletten"
SHEET_EVENTS = "Termine & Dokumentation"
SHEET_PRESCRIPTIONS = "Verordnungen"
SHEET_FINDINGS = "Befunde (Körperschema)"
SHEET_NOTES = "Notizen"
SHEET_FILES = "Dokumente"
QUESTIONNAIRE_PREFIX = "Fragebogen"

# Spalten, mit denen jedes Blatt den Patienten identifiziert
PATIENT_REF_COLS = ["Patienten-Nr", "Nachname", "Vorname", "Geburtsdatum"]

_YES = {"1", "t", "true", "ja", "j", "y", "yes"}
_NO = {"0", "f", "false", "nein", "n", "no"}

PRESCRIPTION_TYPES = {
    "GK": "Gesetzlich (GKV)", "PV": "Privat", "BG": "Berufsgenossenschaft", "HF": "Heilfürsorge",
}
PRESCRIPTION_KINDS = {"BLK": "Blankoverordnung", "ZAE": "Zahnärztliche Verordnung"}
FILE_CATEGORIES = {
    "rezeptvorder": "Rezept Vorderseite",
    "rezeptrueck": "Rezept Rückseite",
    "rezeptanlage": "Rezept Anlage",
}
_VIEW_PARTS = {
    "Male": "männlich", "Female": "weiblich",
    "Muscle": "Muskulatur", "Skeleton": "Skelett",
    "Front": "vorne", "Back": "hinten",
}

# Dubletten-Regeln (Reihenfolge = Stärke; die stärkste Regel einer Gruppe steht vorne)
DUP_RULE_EXACT = "Name, Vorname und Geburtsdatum gleich"
DUP_RULE_SIMILAR = "Gleiches Geburtsdatum, ähnlicher Name"
DUP_RULE_INSURANCE_NO = "Gleiche Versichertennummer"
DUP_RULE_OTHER_BIRTH = "Name und Vorname gleich, anderes Geburtsdatum"
DUP_RULES = [DUP_RULE_EXACT, DUP_RULE_SIMILAR, DUP_RULE_INSURANCE_NO, DUP_RULE_OTHER_BIRTH]
# Mindest-Ähnlichkeit (0..1) des vollständigen Namens für DUP_RULE_SIMILAR
NAME_SIMILARITY_THRESHOLD = 0.85
# Mehr Patienten mit demselben Geburtsdatum deuten auf ein Platzhalter-Datum hin -> kein Ähnlichkeitsvergleich
MAX_SIMILARITY_BLOCK = 200

DATA_PROTECTION_NOTE = (
    "Diese Datei enthält personenbezogene Gesundheitsdaten (Art. 9 DSGVO). "
    "Nur verschlüsselt bzw. über einen sicheren Weg weitergeben."
)


@dataclass
class ExportResult:
    out_path: str
    sheet_counts: dict = field(default_factory=dict)
    missing_tables: list = field(default_factory=list)


# --------------------------------------------------------------------------- #
# Einlesen
# --------------------------------------------------------------------------- #
def find_table_file(folder: str, table: str):
    """Sucht <table>.csv (Groß-/Kleinschreibung egal) im Export-Ordner."""
    wanted = f"{table}.csv".lower()
    for name in os.listdir(folder):
        if name.lower() == wanted and os.path.isfile(os.path.join(folder, name)):
            return os.path.join(folder, name)
    return None


def _clean(df: pd.DataFrame) -> pd.DataFrame:
    """Leerzeichen am Rand entfernen; leere Felder und das Literal "null" werden leer."""
    df = df.copy()
    for col in df.columns:
        if table_io.is_text_column(df[col]):
            s = df[col].str.strip()
            df[col] = s.mask(s.isin(["", "null"]))
    return df


def load_export_folder(folder: str):
    """Liest alle bekannten Tabellen ein. Gibt (tables, missing_tables) zurück."""
    path = find_table_file(folder, REQUIRED_TABLE)
    if path is None:
        raise FileNotFoundError(
            f"Im gewählten Ordner wurde keine Datei '{REQUIRED_TABLE}.csv' gefunden."
        )
    tables = {REQUIRED_TABLE: _clean(table_io.read_table(path))}
    missing = []
    for table in OPTIONAL_TABLES:
        path = find_table_file(folder, table)
        if path is None:
            missing.append(table)
        else:
            tables[table] = _clean(table_io.read_table(path))
    return tables, missing


# --------------------------------------------------------------------------- #
# Wert-Umwandlungen
# --------------------------------------------------------------------------- #
def yes_no(series: pd.Series) -> pd.Series:
    """1/0, t/f, Ja/Nein ... -> "Ja"/"Nein". Unbekannte Werte bleiben unverändert."""
    def conv(v):
        if pd.isna(v):
            return None
        low = str(v).strip().lower()
        if low in _YES:
            return "Ja"
        if low in _NO:
            return "Nein"
        return v
    return series.map(conv).astype(object)


def to_date(series: pd.Series, keep_time: bool = False) -> pd.Series:
    """ISO-Datum(-zeit) -> echtes Excel-Datum. Nicht lesbare Werte bleiben als Text erhalten,
    Daten vor 1900 (kann Excel nicht) werden als deutscher Text ausgegeben."""
    parsed = pd.to_datetime(series, errors="coerce", format="mixed")

    def conv(orig, ts):
        if pd.isna(orig):
            return None
        if pd.isna(ts):
            return orig
        if ts.year < 1900:
            return ts.strftime("%d.%m.%Y %H:%M" if keep_time else "%d.%m.%Y")
        return ts.to_pydatetime().replace(second=0, microsecond=0) if keep_time else ts.date()

    return pd.Series([conv(o, t) for o, t in zip(series, parsed)], index=series.index, dtype=object)


def to_time(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series, errors="coerce", format="mixed")
    return parsed.dt.strftime("%H:%M").astype(object).where(parsed.notna(), None)


def to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def _col(df: pd.DataFrame, name: str) -> pd.Series:
    """Spalte oder leere Serie, falls die Spalte in dieser Programmversion fehlt."""
    if name in df.columns:
        return df[name]
    return pd.Series([None] * len(df), index=df.index, dtype=object)


def format_recipient(value):
    """Rechnungsempfänger: Freitext oder JSON-Liste ([{"anrede", "name", "strassepostfach", "plz", "ort"}])
    -> eine lesbare Zeile, z. B. "Herr Max Muster, Waldweg 1, 12345 Ort"."""
    if pd.isna(value):
        return None
    text = str(value).strip()
    if text.startswith("["):
        try:
            entries = json.loads(text)
        except ValueError:
            entries = None
        if isinstance(entries, list):
            lines = []
            for e in entries:
                if not isinstance(e, dict):
                    continue
                name = " ".join(x for x in [e.get("anrede") if e.get("anrede") in ("Herr", "Frau") else None,
                                            e.get("name")] if x)
                city = " ".join(x for x in [e.get("plz"), e.get("ort")] if x)
                lines.append(", ".join(x for x in [name, e.get("strassepostfach"), city] if x))
            return "; ".join(l for l in lines if l) or None
    return ", ".join(p.strip() for p in text.splitlines() if p.strip()) or None


def view_label(name) -> str:
    """'Female_Muscle_Front' -> 'Muskulatur vorne (weiblich)'."""
    if pd.isna(name):
        return None
    parts = str(name).split("_")
    if len(parts) == 3 and all(p in _VIEW_PARTS for p in parts):
        gender, layer, side = (_VIEW_PARTS[p] for p in parts)
        return f"{layer} {side} ({gender})"
    return name


def sheet_name(name: str, used: set) -> str:
    """Gültiger, eindeutiger Excel-Blattname (max. 31 Zeichen, ohne []:*?/\\)."""
    base = re.sub(r"[\[\]:*?/\\]", " ", name).strip()[:31] or "Blatt"
    candidate, n = base, 2
    while candidate.lower() in {u.lower() for u in used}:
        suffix = f" ({n})"
        candidate = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


# --------------------------------------------------------------------------- #
# Patientenbezug
# --------------------------------------------------------------------------- #
def patient_reference(patients: pd.DataFrame) -> pd.DataFrame:
    """Tabelle interne Patienten-ID -> Anzeige-Spalten (PATIENT_REF_COLS) + Sortierschlüssel."""
    ref = pd.DataFrame({
        "__pid": patients["id"],
        "Patienten-Nr": _col(patients, "p_nr").fillna(patients["id"]),
        "Nachname": _col(patients, "p_name"),
        "Vorname": _col(patients, "p_vname"),
        "Geburtsdatum": to_date(_col(patients, "p_birth")),
    })
    return ref.drop_duplicates("__pid")


def attach_patient(df: pd.DataFrame, pid: pd.Series, ref: pd.DataFrame) -> pd.DataFrame:
    """Stellt die Patienten-Spalten vor df. Unbekannte Patienten behalten ihre ID als Patienten-Nr."""
    keys = pd.DataFrame({"__pid": pid.values})
    merged = keys.merge(ref, on="__pid", how="left")
    merged["Patienten-Nr"] = merged["Patienten-Nr"].fillna(merged["__pid"])
    merged.index = df.index
    return pd.concat([merged[PATIENT_REF_COLS], df], axis=1)


_UMLAUTS = str.maketrans({"ä": "a", "ö": "o", "ü": "u", "ß": "ss"})


def _german_key(series: pd.Series) -> pd.Series:
    """Sortierschlüssel nach deutscher Telefonbuch-Logik (Umlaute wie Grundbuchstaben)."""
    return series.fillna("").astype(str).str.lower().map(lambda v: v.translate(_UMLAUTS))


def sort_by_patient(df: pd.DataFrame, *extra_raw_keys: pd.Series) -> pd.DataFrame:
    """Sortiert nach Nachname, Vorname, Patienten-Nr und optional weiteren (Roh-)Schlüsseln."""
    keys = pd.DataFrame({
        "a": _german_key(df["Nachname"]),
        "b": _german_key(df["Vorname"]),
        "c": df["Patienten-Nr"].fillna("").astype(str),
    }, index=df.index)
    for i, k in enumerate(extra_raw_keys):
        keys[f"k{i}"] = k.fillna("").astype(str).values
    order = keys.sort_values(list(keys.columns), kind="stable").index
    return df.loc[order].reset_index(drop=True)


class _RefResolver:
    """Löst (reftable, refid) aus history/files zum Patienten + lesbarem Bezug auf."""

    def __init__(self, tables):
        self.prescriptions = {}
        if "rezepte" in tables:
            z = tables["rezepte"]
            self.prescriptions = dict(zip(z["id"], zip(z["patienten_id"], _col(z, "rezeptnr"))))
        self.events = {}
        if "events" in tables:
            e = tables["events"]
            self.events = dict(zip(e["id"], zip(e["patient_id"], _col(e, "start"))))

    def resolve(self, reftable, refid):
        if reftable == "patienten":
            return refid, "Patient"
        if reftable == "rezepte" and refid in self.prescriptions:
            pid, nr = self.prescriptions[refid]
            return pid, f"Verordnung {nr}" if pd.notna(nr) else "Verordnung"
        if reftable == "events" and refid in self.events:
            pid, start = self.events[refid]
            parsed = pd.to_datetime(start, errors="coerce")
            return pid, f"Termin {parsed:%d.%m.%Y}" if pd.notna(parsed) else "Termin"
        return None, None

    def apply(self, df: pd.DataFrame):
        resolved = [self.resolve(t, r) for t, r in zip(df["reftable"], df["refid"])]
        pid = pd.Series([r[0] for r in resolved], index=df.index, dtype=object)
        bezug = pd.Series([r[1] for r in resolved], index=df.index, dtype=object)
        return pid, bezug


# --------------------------------------------------------------------------- #
# Blätter
# --------------------------------------------------------------------------- #
def build_patients(t) -> pd.DataFrame:
    p = t["patienten"]
    df = pd.DataFrame({
        "Patienten-Nr": _col(p, "p_nr").fillna(p["id"]),
        "Anrede": _col(p, "p_anrede"),
        "Nachname": _col(p, "p_name"),
        "Vorname": _col(p, "p_vname"),
        "Geburtsdatum": to_date(_col(p, "p_birth")),
        "Straße": _col(p, "p_street"),
        "Hausnummer": _col(p, "p_hausnummer"),
        "PLZ": _col(p, "p_plz"),
        "Ort": _col(p, "p_ort"),
        "Telefon": _col(p, "p_tel"),
        "Telefon geschäftlich": _col(p, "p_telge"),
        "Mobil": _col(p, "p_handy"),
        "E-Mail": _col(p, "p_email"),
        "Krankenkasse": _col(p, "p_krankenkasse"),
        "IK Krankenkasse": _col(p, "p_ik"),
        "Versichertennummer": _col(p, "p_vnr"),
        "Versichertenstatus": _col(p, "p_vs"),
        "Privatversichert": yes_no(_col(p, "p_privatversichert")),
        "Zuzahlungsbefreit": yes_no(_col(p, "p_zuzahlungsbefreit")),
        "Zuzahlungsbefreit bis": to_date(_col(p, "p_zuzahlungsbefreit_bis")),
        "Abweichender Rechnungsempfänger": _col(p, "rechnungsempfaenger").map(format_recipient).astype(object),
        "VIP": yes_no(_col(p, "p_vip")),
        "Angelegt am": to_date(_col(p, "erstellt_am")),
    })
    return sort_by_patient(df)


def build_events(t, ref) -> pd.DataFrame:
    if "events" not in t:
        return None
    e = t["events"]
    e = e[e["patient_id"].notna()]

    therapists = {}
    if "resources" in t:
        therapists = dict(zip(t["resources"]["id"], t["resources"]["name"]))

    prescription_of_event = {}
    if "rezepte_events" in t and "rezepte" in t:
        numbers = dict(zip(t["rezepte"]["id"], _col(t["rezepte"], "rezeptnr")))
        re_ = t["rezepte_events"].dropna(subset=["events_id"])
        nr = re_["rezepte_id"].map(numbers).fillna(re_["rezepte_id"])
        prescription_of_event = nr.groupby(re_["events_id"]).agg(
            lambda s: ", ".join(sorted(set(s.astype(str))))
        ).to_dict()

    df = pd.DataFrame({
        "Datum": to_date(_col(e, "start")),
        "Beginn": to_time(_col(e, "start")),
        "Ende": to_time(_col(e, "end")),
        "Dauer (Min.)": to_number(_col(e, "duration")),
        "Behandlung": _col(e, "title"),
        "Therapeut": _col(e, "resource").map(lambda r: therapists.get(r, r) if pd.notna(r) else None),
        "Verordnung (Rezept-Nr)": e["id"].map(prescription_of_event),
        "Termin der Serie": _col(e, "eventsequenz"),
        "Wahrgenommen": yes_no(_col(e, "wahrgenommen")),
        "Nicht erschienen": yes_no(_col(e, "nicht_wahrgenommen")),
        "Abgesagt": yes_no(_col(e, "abgesagt")),
        "Hausbesuch": yes_no(_col(e, "hausbesuch")),
        "Abgerechnet": yes_no(_col(e, "abgerechnet")),
        "Stichwort": _col(e, "dok_stichwort"),
        "Dokumentation": _col(e, "dokumentation"),
        "Bemerkung": _col(e, "description"),
        "Memo": _col(e, "memo"),
    }, index=e.index)
    df = attach_patient(df, e["patient_id"], ref)
    return sort_by_patient(df, _col(e, "start"))


def _remedies(z: pd.DataFrame, pairs) -> pd.Series:
    """Heilmittel + Einheiten mehrerer Spalten zu "KG (6x); MT (10x)" zusammenfassen."""
    def row_text(row):
        parts = []
        for name_col, units_col in pairs:
            name, units = row.get(name_col), row.get(units_col)
            if pd.notna(name):
                parts.append(f"{name} ({units}x)" if pd.notna(units) else str(name))
        return "; ".join(parts) or None
    return z.apply(row_text, axis=1).astype(object) if len(z) else pd.Series(dtype=object)


def build_prescriptions(t, ref) -> pd.DataFrame:
    if "rezepte" not in t:
        return None
    z = t["rezepte"]
    event_count = pd.Series(dtype=float)
    if "rezepte_events" in t:
        re_ = t["rezepte_events"].dropna(subset=["events_id"])
        event_count = re_.groupby("rezepte_id")["events_id"].nunique()

    df = pd.DataFrame({
        "Rezept-Nr": _col(z, "rezeptnr"),
        "Ausstellungsdatum": to_date(_col(z, "datum")),
        "Kostenträgerart": _col(z, "rezept_typ").map(lambda v: PRESCRIPTION_TYPES.get(v, v)),
        "Verordnungsart": _col(z, "verordnungstyp").map(lambda v: PRESCRIPTION_KINDS.get(v, v)),
        "Kostenträger": _col(z, "krankenkasse"),
        "Versichertennummer": _col(z, "versichertennr"),
        "Versichertenstatus": _col(z, "status"),
        "Arzt-Nr (LANR)": _col(z, "arztnr"),
        "Betriebsstätten-Nr (BSNR)": _col(z, "betriebsstaettennr"),
        "Diagnosegruppe": _col(z, "diagnosegruppe"),
        "ICD-10 (1)": _col(z, "icd10_1"),
        "ICD-10 (2)": _col(z, "icd10_2"),
        "Diagnose": _col(z, "diagnose"),
        "Leitsymptomatik": _col(z, "leitsymptomatik_text").str.rstrip(" ;"),
        "Heilmittel": _remedies(z, [("vorrangiges_heilmittel", "behandlungseinheiten"),
                                    ("vorrangiges_heilmittel2", "behandlungseinheiten2"),
                                    ("vorrangiges_heilmittel3", "behandlungseinheiten3")]),
        "Ergänzendes Heilmittel": _remedies(z, [("ergaenzendes_heilmittel",
                                                 "ergaenzende_behandlungseinheiten")]),
        "Therapiefrequenz": _col(z, "therapiefrequenz"),
        "Hausbesuch": yes_no(_col(z, "hausbesuch")),
        "Dringlicher Behandlungsbedarf": yes_no(_col(z, "dringlichkeit")),
        "Therapiebericht angefordert": yes_no(_col(z, "therapiebericht")),
        "Behandlungsbeginn spätestens": to_date(_col(z, "behandlungsanfang_max")),
        "Behandlungsende spätestens": to_date(_col(z, "behandlungsende_max")),
        "Anzahl Termine": z["id"].map(event_count),
        "Erledigt": yes_no(_col(z, "rezept_erledigt")),
        "Abgerechnet am": to_date(_col(z, "rezept_abgerechnet_am")),
        "Weitere Befunde / Therapieziele": _col(z, "weitere_befunde"),
        "Therapieerfolg": _col(z, "tp_erfolg"),
        "Therapiebericht erstellt am": to_date(_col(z, "tp_erstellt_am")),
        "Therapiebericht": _col(z, "tp_therapiebericht"),
    }, index=z.index)
    df = attach_patient(df, z["patienten_id"], ref)
    return sort_by_patient(df, _col(z, "datum"))


def build_findings(t, ref) -> pd.DataFrame:
    if "anatomy_patienten" not in t:
        return None
    a = t["anatomy_patienten"]
    views = {}
    if "anatomy" in t:
        views = dict(zip(t["anatomy"]["id"], t["anatomy"]["name"].map(view_label)))
    region = _col(a, "id_region").map(views)
    if "name_region" in a.columns:
        region = a["name_region"].fillna(region)
    df = pd.DataFrame({
        "Datum": to_date(_col(a, "timestamp")),
        "Ansicht": region,
        "Befund": _col(a, "title"),
        "Notiz": _col(a, "notes"),
    }, index=a.index)
    df = attach_patient(df, a["id_patient"], ref)
    return sort_by_patient(df, _col(a, "timestamp"))


def build_questionnaires(t, ref) -> dict:
    """Ein Blatt je Fragebogen (Kartei): eine Zeile je ausgefülltem Bogen, eine Spalte je Frage."""
    needed = ("kartei_history", "kartei_content", "pat_property")
    if any(n not in t for n in needed) or t["kartei_history"].empty:
        return {}
    hist, content, props = (t[n] for n in needed)
    names = {}
    if "kartei" in t:
        names = dict(zip(t["kartei"]["id"], t["kartei"]["karteiname"]))

    result = {}
    for kartei_id, h in hist.groupby("kartei_id", sort=False):
        kprops = props[props["kartei_id"] == kartei_id].copy()
        kprops["__sort"] = to_number(_col(kprops, "sortierung")).fillna(0)
        kprops = kprops.sort_values("__sort", kind="stable")

        # Eindeutige Spaltennamen je Frage
        labels, seen = {}, {}
        for pid, label in zip(kprops["id"], kprops["label"].fillna("Frage")):
            seen[label] = seen.get(label, 0) + 1
            labels[pid] = label if seen[label] == 1 else f"{label} ({seen[label]})"
        types = dict(zip(kprops["id"], _col(kprops, "proptyp")))

        c = content[content["kartei_history_id"].isin(h["id"])
                    & content["kartei_property_id"].isin(labels)]
        answers = {}
        for hid, prop, val in zip(c["kartei_history_id"], c["kartei_property_id"], c["content"]):
            if pd.isna(val):
                continue
            if types.get(prop) == "CHK":
                val = yes_no(pd.Series([val])).iloc[0]
            answers.setdefault(hid, {})[labels[prop]] = val

        df = pd.DataFrame({
            "Ausgefüllt am": to_date(_col(h, "datum")),
            "Gültig bis": to_date(_col(h, "expires")),
        }, index=h.index)
        for label in labels.values():
            df[label] = h["id"].map(lambda hid, l=label: answers.get(hid, {}).get(l))
        df = attach_patient(df, h["patienten_id"], ref)
        title = names.get(kartei_id)
        title = f"{QUESTIONNAIRE_PREFIX} {title}" if pd.notna(title) and title else QUESTIONNAIRE_PREFIX
        result[title] = sort_by_patient(df, _col(h, "datum"))
    return result


def build_notes(t, ref, resolver) -> pd.DataFrame:
    if "history" not in t:
        return None
    h = t["history"]
    pid, bezug = resolver.apply(h)
    keep = pid.notna()
    h, pid, bezug = h[keep], pid[keep], bezug[keep]
    df = pd.DataFrame({
        "Datum": to_date(_col(h, "datum"), keep_time=True),
        "Bezug": bezug,
        "Stichwort": _col(h, "stichwort"),
        "Text": _col(h, "text"),
        "Bearbeiter": _col(h, "bearbeiter"),
    }, index=h.index)
    df = attach_patient(df, pid, ref)
    return sort_by_patient(df, _col(h, "datum"))


def build_files(t, ref, resolver) -> pd.DataFrame:
    if "files" not in t:
        return None
    f = t["files"]
    pid, bezug = resolver.apply(f)
    keep = pid.notna()
    f, pid, bezug = f[keep], pid[keep], bezug[keep]
    size_kb = (to_number(_col(f, "size")) / 1024).round(0)
    df = pd.DataFrame({
        "Datum": to_date(_col(f, "erstellt_am")),
        "Bezug": bezug,
        "Kategorie": _col(f, "kategorie").map(lambda v: FILE_CATEGORIES.get(v, v)),
        "Dateiname": _col(f, "orgname"),
        "Beschreibung": _col(f, "beschreibung"),
        "Dateityp": _col(f, "ext").str.upper(),
        "Größe (KB)": size_kb,
        "Ablagename": _col(f, "filename"),
    }, index=f.index)
    df = attach_patient(df, pid, ref)
    return sort_by_patient(df, _col(f, "erstellt_am"))


def build_sheets(tables: dict) -> dict:
    """Erzeugt alle Blätter (Name -> DataFrame) in Ausgabe-Reihenfolge, ohne Übersicht.
    Blätter, deren Quelltabellen fehlen, werden ausgelassen."""
    ref = patient_reference(tables["patienten"])
    resolver = _RefResolver(tables)
    sheets = {
        SHEET_PATIENTS: build_patients(tables),
        SHEET_DUPLICATES: build_duplicates(tables, resolver),
    }
    candidates = [
        (SHEET_EVENTS, build_events(tables, ref)),
        (SHEET_PRESCRIPTIONS, build_prescriptions(tables, ref)),
        (SHEET_FINDINGS, build_findings(tables, ref)),
    ]
    candidates += list(build_questionnaires(tables, ref).items())
    candidates += [
        (SHEET_NOTES, build_notes(tables, ref, resolver)),
        (SHEET_FILES, build_files(tables, ref, resolver)),
    ]
    for name, df in candidates:
        if df is not None:
            sheets[name] = df
    return sheets


# --------------------------------------------------------------------------- #
# Dubletten
# --------------------------------------------------------------------------- #
_NAME_TRANSLATION = str.maketrans({"ä": "ae", "ö": "oe", "ü": "ue", "ß": "ss"})


def normalize_name(value) -> str:
    """Für den Vergleich: klein, Umlaute ausgeschrieben, Akzente entfernt, nur Buchstaben.
    "Müller-Lüdenscheidt " -> "muellerluedenscheidt", "Ayşe" -> "ayse"."""
    if pd.isna(value):
        return ""
    text = str(value).lower().translate(_NAME_TRANSLATION)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"[^a-z]", "", text)


def normalize_insurance_no(value) -> str:
    """Versichertennummer ohne Leer-/Trennzeichen; Platzhalter (z. B. 000000) zählen als leer."""
    if pd.isna(value):
        return ""
    text = re.sub(r"[^0-9A-Za-z]", "", str(value)).upper()
    if len(text) < 6 or len(set(text)) == 1:
        return ""
    return text


def name_similarity(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, a, b).ratio()


def find_duplicate_pairs(patients: pd.DataFrame) -> list:
    """Liefert (Index_a, Index_b, Regel)-Paare möglicher Dubletten.

    Gleichheits-Regeln werden als Kette verknüpft (A-B, B-C statt aller Paare), damit auch
    viele gleichartige Datensätze (z. B. hunderte "Test, Test") linear bleiben. Der
    Ähnlichkeitsvergleich läuft nur innerhalb eines Geburtsdatums; auffällig große Blöcke
    (Platzhalter-Datum wie 01.01.1900) werden dabei übersprungen."""
    last = _col(patients, "p_name").map(normalize_name)
    first = _col(patients, "p_vname").map(normalize_name)
    birth = _col(patients, "p_birth").fillna("").astype(str)
    insurance_no = _col(patients, "p_vnr").map(normalize_insurance_no)
    has_name = (last != "") & (first != "")
    pairs = []

    def chain(index, rule):
        index = list(index)
        pairs.extend((a, b, rule) for a, b in zip(index, index[1:]))

    # 1 + 4: gleicher (normalisierter) Name und Vorname
    names = pd.DataFrame({"l": last, "f": first, "b": birth})[has_name]
    for _, group in names.groupby(["l", "f"], sort=False):
        if len(group) < 2:
            continue
        representatives = []
        for b_value, same_birth in group.groupby("b", sort=False):
            if b_value:
                chain(same_birth.index, DUP_RULE_EXACT)
                representatives.append(same_birth.index[0])
            else:  # ohne Geburtsdatum kann nur "anderes Geburtsdatum" gelten
                representatives.extend(same_birth.index)
        chain(representatives, DUP_RULE_OTHER_BIRTH)

    # 2: gleiches Geburtsdatum, ähnlicher (nicht identischer) Name – auch Vor-/Nachname vertauscht
    full = last + " " + first
    swapped = first + " " + last
    candidates = pd.DataFrame({"b": birth, "full": full})[(birth != "") & has_name]
    for _, group in candidates.groupby("b", sort=False):
        if len(group) < 2 or len(group) > MAX_SIMILARITY_BLOCK:
            continue
        unique_names = group.drop_duplicates("full").index.tolist()  # identische = schon Regel 1
        for i, a in enumerate(unique_names):
            for b in unique_names[i + 1:]:
                score = max(name_similarity(full[a], full[b]), name_similarity(full[a], swapped[b]))
                if score >= NAME_SIMILARITY_THRESHOLD:
                    pairs.append((a, b, DUP_RULE_SIMILAR))

    # 3: gleiche Versichertennummer
    for _, group in pd.DataFrame({"v": insurance_no})[insurance_no != ""].groupby("v", sort=False):
        if len(group) > 1:
            chain(group.index, DUP_RULE_INSURANCE_NO)
    return pairs


def _group_pairs(pairs: list) -> list:
    """Verbindet Paare zu Gruppen (A~B, B~C -> A,B,C). Gibt [(Mitglieder, Regeln)] zurück."""
    parent = {}

    def find(x):
        parent.setdefault(x, x)
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for a, b, _ in pairs:
        parent[find(a)] = find(b)
    groups = {}
    for a, b, rule in pairs:
        members, rules = groups.setdefault(find(a), (set(), set()))
        members.update((a, b))
        rules.add(rule)
    return list(groups.values())


def build_duplicates(t, resolver=None) -> pd.DataFrame:
    """Blatt "Mögliche Dubletten": eine Zeile je Patient, zusammengehörige Datensätze tragen
    dieselbe Gruppen-Nr. Zusatzspalten (Termine, Verordnungen ...) helfen bei der Entscheidung,
    welcher Datensatz behalten werden soll."""
    p = t["patienten"].drop_duplicates("id").reset_index(drop=True)
    columns = ["Gruppe", "Übereinstimmung"] + PATIENT_REF_COLS + [
        "Anrede", "Anschrift", "Telefon", "Krankenkasse", "Versichertennummer", "Angelegt am",
        "Anzahl Termine", "Letzter Termin", "Anzahl Verordnungen", "Anzahl Dokumente",
    ]
    groups = _group_pairs(find_duplicate_pairs(p))
    if not groups:
        return pd.DataFrame(columns=columns)

    event_count, last_event, rx_count, file_count = {}, {}, {}, {}
    if "events" in t:
        e = t["events"].dropna(subset=["patient_id"])
        event_count = e.groupby("patient_id")["id"].count().to_dict()
        last_event = e.groupby("patient_id")["start"].max().to_dict() if "start" in e else {}
    if "rezepte" in t:
        rx_count = t["rezepte"].groupby("patienten_id")["id"].count().to_dict()
    if "files" in t and resolver is not None:
        pid, _ = resolver.apply(t["files"])
        file_count = pid.dropna().value_counts().to_dict()

    rule_rank = {rule: i for i, rule in enumerate(DUP_RULES)}
    sort_name = _german_key(_col(p, "p_name")) + " " + _german_key(_col(p, "p_vname"))
    ordered = sorted(
        groups,
        key=lambda g: (min(rule_rank[r] for r in g[1]), min(sort_name[i] for i in g[0])),
    )

    rows = []
    for number, (members, rules) in enumerate(ordered, start=1):
        reason = "; ".join(sorted(rules, key=rule_rank.get))
        for i in sorted(members, key=lambda i: (sort_name[i], str(_col(p, "p_birth")[i]))):
            r = p.loc[i]
            pid = r["id"]
            street = " ".join(str(x) for x in [r.get("p_street"), r.get("p_hausnummer")] if pd.notna(x))
            city = " ".join(str(x) for x in [r.get("p_plz"), r.get("p_ort")] if pd.notna(x))
            phone = r.get("p_tel") if pd.notna(r.get("p_tel")) else r.get("p_handy")
            rows.append({
                "Gruppe": number,
                "Übereinstimmung": reason,
                "Patienten-Nr": r.get("p_nr") if pd.notna(r.get("p_nr")) else pid,
                "Nachname": r.get("p_name"),
                "Vorname": r.get("p_vname"),
                "Geburtsdatum": r.get("p_birth"),
                "Anrede": r.get("p_anrede"),
                "Anschrift": ", ".join(x for x in [street, city] if x) or None,
                "Telefon": phone,
                "Krankenkasse": r.get("p_krankenkasse"),
                "Versichertennummer": r.get("p_vnr"),
                "Angelegt am": r.get("erstellt_am"),
                "Anzahl Termine": event_count.get(pid, 0),
                "Letzter Termin": last_event.get(pid),
                "Anzahl Verordnungen": rx_count.get(pid, 0),
                "Anzahl Dokumente": file_count.get(pid, 0),
            })
    df = pd.DataFrame(rows, columns=columns)
    for col in ("Geburtsdatum", "Angelegt am", "Letzter Termin"):
        df[col] = to_date(df[col].astype(object))
    return df


def build_overview(sheets: dict, missing: list, source: str) -> pd.DataFrame:
    """Deckblatt: Erstellungszeitpunkt, Quelle, Datenschutzhinweis und Anzahl Einträge je Blatt."""
    rows = [
        ("Erstellt am", f"{datetime.datetime.now():%d.%m.%Y %H:%M}"),
        ("Quelle (Export-Ordner)", source),
        ("Datenschutz", DATA_PROTECTION_NOTE),
        (None, None),
    ]
    rows += [(f"Blatt \u201e{name}\u201c", f"{len(df)} Eintrag" if len(df) == 1 else f"{len(df)} Einträge") for name, df in sheets.items()]
    if missing:
        rows += [(None, None), ("Nicht im Export enthalten", ", ".join(f"{m}.csv" for m in missing))]
    return pd.DataFrame(rows, columns=["Angabe", "Wert"])


def write_workbook(sheets: dict, overview: pd.DataFrame, out_path: str) -> dict:
    """Schreibt die Arbeitsmappe. Gibt die tatsächlich verwendeten Blattnamen -> Zeilen zurück."""
    used, counts = set(), {}
    with table_io.excel_writer(out_path, date_format="DD.MM.YYYY",
                               datetime_format="DD.MM.YYYY HH:MM") as writer:
        table_io.write_excel_sheet(writer, overview, sheet_name(SHEET_OVERVIEW, used), autofilter=False)
        for name, df in sheets.items():
            final = sheet_name(name, used)
            table_io.write_excel_sheet(writer, df, final)
            counts[final] = len(df)
    return counts


def export_patient_record(folder: str, out_path: str) -> ExportResult:
    """Komplettablauf: Export-Ordner lesen -> Blätter bauen -> Excel schreiben."""
    tables, missing = load_export_folder(folder)
    sheets = build_sheets(tables)
    overview = build_overview(sheets, missing, os.path.basename(os.path.normpath(folder)))
    counts = write_workbook(sheets, overview, out_path)
    return ExportResult(out_path=out_path, sheet_counts=counts, missing_tables=missing)

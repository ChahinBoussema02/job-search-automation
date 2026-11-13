#!/usr/bin/env python3
"""
jobfinder.py — findet passende Stellen über die offizielle Jobsuche-API der
Bundesagentur für Arbeit und ergänzt neue Treffer in einer Excel-Datei.

Es bewirbt sich nie. Es sucht, filtert, entdoppelt und protokolliert.
Das Absenden bleibt Handarbeit.

Benutzung
---------
    pip install requests openpyxl
    python3 jobfinder.py                      # normaler Lauf
    python3 jobfinder.py --debug              # Rohfelder des ersten Treffers
    python3 jobfinder.py --out Tracker.xlsx   # andere Zieldatei

API: /pc/v6/jobs — v4 wurde abgeschaltet und antwortet mit 403.
Doku: https://github.com/bundesAPI/jobsuche-api
"""

import argparse
import datetime as dt
import re
import sys
import time
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

API = "https://rest.arbeitsagentur.de/jobboerse/jobsuche-service/pc/v6/jobs"
HEADERS = {
    "X-API-Key": "jobboerse-jobsuche",
    "Accept": "application/json",
    "User-Agent": "jobfinder/2.0",
}
DETAIL_URL = "https://www.arbeitsagentur.de/jobsuche/jobdetail/{ref}"

# ---------------------------------------------------------------------------
# Suchprofile. Jedes Dict wird eine Abfrage.
#   was      Suchbegriff
#   wo       Ort (weglassen = bundesweit)
#   umkreis  Radius in km um `wo`
#   remote   True = bundesweit suchen und danach auf Homeoffice filtern
#            (v6 hat keinen funktionierenden Serverfilter dafür)
# Diese Liste ist die gesamte Konfiguration — hier erweitern.
# ---------------------------------------------------------------------------
SEARCHES = [
    # --- Pflichtpraktikum / Praktikum, Umgebung Krefeld -------------------
    {"label": "Pflichtpraktikum Software NRW",
     "was": "Pflichtpraktikum Softwareentwicklung", "wo": "Krefeld", "umkreis": 80},
    {"label": "Pflichtpraktikum Informatik NRW",
     "was": "Pflichtpraktikum Informatik", "wo": "Krefeld", "umkreis": 80},
    {"label": "Praktikum Softwareentwicklung NRW",
     "was": "Praktikum Softwareentwicklung", "wo": "Krefeld", "umkreis": 80},
    {"label": "Praktikum Informatik NRW",
     "was": "Praktikum Informatik", "wo": "Krefeld", "umkreis": 80},
    {"label": "Praktikum KI NRW",
     "was": "Praktikum Künstliche Intelligenz", "wo": "Krefeld", "umkreis": 80},
    {"label": "Praktikum Machine Learning NRW",
     "was": "Praktikum Machine Learning", "wo": "Krefeld", "umkreis": 80},
    {"label": "Praktikum Data Science NRW",
     "was": "Praktikum Data Science", "wo": "Krefeld", "umkreis": 80},

    # --- Werkstudent, Umgebung Krefeld ------------------------------------
    {"label": "Werkstudent Softwareentwicklung NRW",
     "was": "Werkstudent Softwareentwicklung", "wo": "Krefeld", "umkreis": 60},
    {"label": "Werkstudent Informatik NRW",
     "was": "Werkstudent Informatik", "wo": "Krefeld", "umkreis": 60},
    {"label": "Werkstudent Fullstack NRW",
     "was": "Werkstudent Fullstack Entwickler", "wo": "Krefeld", "umkreis": 60},
    {"label": "Werkstudent Backend NRW",
     "was": "Werkstudent Backend Entwickler", "wo": "Krefeld", "umkreis": 60},
    {"label": "Werkstudent IT NRW",
     "was": "Werkstudent IT", "wo": "Krefeld", "umkreis": 60},
    {"label": "Werkstudent KI NRW",
     "was": "Werkstudent Künstliche Intelligenz", "wo": "Krefeld", "umkreis": 80},
    {"label": "Werkstudent Data Science NRW",
     "was": "Werkstudent Data Science", "wo": "Krefeld", "umkreis": 80},
    {"label": "Werkstudent Machine Learning NRW",
     "was": "Werkstudent Machine Learning", "wo": "Krefeld", "umkreis": 80},
    {"label": "Werkstudent Automatisierung NRW",
     "was": "Werkstudent Prozessautomatisierung", "wo": "Krefeld", "umkreis": 80},
    {"label": "Werkstudent DevOps NRW",
     "was": "Werkstudent DevOps", "wo": "Krefeld", "umkreis": 60},

    # --- Bundesweit, danach auf Homeoffice gefiltert -----------------------
    {"label": "Praktikum Software remote",
     "was": "Praktikum Softwareentwicklung", "remote": True},
    {"label": "Praktikum KI remote",
     "was": "Praktikum Künstliche Intelligenz", "remote": True},
    {"label": "Praktikum Machine Learning remote",
     "was": "Praktikum Machine Learning", "remote": True},
    {"label": "Praktikum Data Science remote",
     "was": "Praktikum Data Science", "remote": True},
    {"label": "Werkstudent KI remote",
     "was": "Werkstudent Künstliche Intelligenz", "remote": True},
    {"label": "Werkstudent Data Science remote",
     "was": "Werkstudent Data Science", "remote": True},
    {"label": "Werkstudent Software remote",
     "was": "Werkstudent Softwareentwicklung", "remote": True},
]

MAX_AGE_DAYS = 60      # v6 ignoriert veroeffentlichtseit — wir filtern selbst
PAGE_SIZE = 100
MAX_PAGES = 3

# stellenangebotsart lesbar machen
ARTEN = {
    "PRAKTIKUM_TRAINEE": "Praktikum/Trainee",
    "ARBEIT": "Arbeit",
    "AUSBILDUNG": "Ausbildung",
    "SELBSTAENDIGKEIT": "Selbstständigkeit",
    "BUFDI_FSJ_FOEJ": "Freiwilligendienst",
}
ARTEN_BLOCKLIST = {"AUSBILDUNG", "SELBSTAENDIGKEIT", "BUFDI_FSJ_FOEJ"}

# Titel mit einem dieser Wörter werden verworfen. Spart Sortierarbeit.
TITLE_BLOCKLIST = [
    "ausbildung", "schülerpraktikum", "schulpraktikum", "schüler",
    "vertrieb", "sales", "pflege", "handwerk", "reinigung", "zerspanung",
    "verkäufer", "lagerhelfer", "produktionshelfer", "kraftfahrer",
    "einkauf", "buchhaltung", "steuerberat", "lohnbuchhalt", "recruiting",
    "personalreferent", "immobilien", "versicherungsberat",
]

# Nur Stellen, die auch wirklich für Studierende sind. Die API nimmt
# "Werkstudent" nur als Stichwort, nicht als Bedingung — ohne diesen Filter
# kommen unbefristete Vollzeitstellen mit durch.
STUDENT_RX = re.compile("|".join([
    r"werkstudent", r"praktik", r"studentisch", r"working student",
    r"\bintern\b", r"\binternship\b", r"abschlussarbeit", r"bachelorarbeit",
    r"masterarbeit", r"thesis", r"\bstudent", r"\bwerkstudi",
]), re.I)

# Allgemein technische Begriffe — für die Relevanzspalte, nicht zum Filtern.
TECH_RX = re.compile("|".join([
    r"informatik", r"software", r"entwickl", r"programmier", r"\bit\b",
    r"\bit-", r"\biot\b", r"cloud", r"devops", r"digitalisierung",
    r"automatisier", r"cyber", r"security", r"embedded", r"datenbank",
    r"netzwerk", r"\bdaten\b", r"\bdata\b", r"system", r"\btest",
    r"\bqa\b", r"robotik", r"simulation", r"\bweb\b", r"\bapp\b",
    r"engineer", r"technolog", r"digital",
]), re.I)

# --- Bewertung -------------------------------------------------------------
# Punkte, die ein Treffer im Stellentitel bringt. Hoch gewichtet ist, was im
# Lebenslauf mit Projekterfahrung hinterlegt ist; niedrig, was nur ein
# technisches Umfeld andeutet. Jedes Muster zählt einmal, nicht pro Fundstelle.
# Hier schrauben, wenn sich der Schwerpunkt verschiebt.
TECH_PUNKTE = [
    # Kerntechnologien aus dem Lebenslauf
    (r"\bpython\b", 10), (r"\bpytorch\b", 10), (r"\btypescript\b", 9),
    (r"\breact\b", 9), (r"next\.?js", 9), (r"nest\.?js", 9),
    (r"node\.?js", 8), (r"\bc#|c-sharp", 8), (r"\.net\b", 8),
    (r"fastapi", 8), (r"postgres", 7), (r"javascript", 6),
    (r"\bdocker\b", 6), (r"\bangular\b", 5), (r"\bsql\b", 5),
    (r"\bapi\b|\brest\b", 5), (r"scikit|tensorflow", 7),

    # Schwerpunkt KI
    (r"\bllm\b|sprachmodell|large language", 12),
    (r"\brag\b|retrieval.{0,3}aug", 12),
    (r"künstliche intelligenz|\bki\b|\bai\b|artificial intelligence", 12),
    (r"machine learning|\bml\b|maschinelles lernen", 10),
    (r"deep learning|neuronale", 9),
    (r"\bnlp\b|natural language", 9),
    (r"generativ|genai", 8),
    (r"data scien|datenanaly|data analy", 6),

    # Breitere Entwicklungsbegriffe
    (r"full ?-?stack", 8), (r"backend|back-end", 6),
    (r"entwickl|developer|programmier", 6),
    (r"software", 5), (r"frontend|front-end", 5),
    (r"\bweb\b|webentwickl", 4), (r"\bdevops\b", 4),
    (r"informatik", 4), (r"automatisier", 4), (r"\bcloud\b", 3),

    # Nur technisches Umfeld — reicht allein nicht für "hoch"
    (r"embedded", 4), (r"security|cyber", 4), (r"\bit\b|\bit-", 4),
    (r"\bdaten\b|\bdata\b", 3), (r"\btest|\bqa\b", 3), (r"\biot\b", 3),
    (r"datenbank", 3), (r"robotik|simulation", 3), (r"netzwerk", 2),
    (r"system", 2), (r"technolog|digital", 2),
]
TECH_PUNKTE = [(re.compile(p, re.I), n) for p, n in TECH_PUNKTE]

TECH_DECKEL = 40        # Titel mit Schlagwortkette nicht davonlaufen lassen
PFLICHT_RX = re.compile(r"pflichtpraktikum|pflicht-praktikum", re.I)

# Grenzen für die Spalte "Relevanz"
SCHWELLE_HOCH = 16
SCHWELLE_MITTEL = 8

# --- CV-Auswahl ------------------------------------------------------------
CV_KI = "Lebenslauf-KI.pdf"
CV_SW = "Lebenslauf-Software.pdf"

KI_RX = re.compile("|".join([
    r"\bki\b", r"künstliche intelligenz", r"\bai\b", r"artificial intelligence",
    r"machine learning", r"\bml\b", r"data scien", r"\bllm\b", r"sprachmodell",
    r"\bnlp\b", r"deep learning", r"neuronale", r"generativ", r"data analy",
    r"datenanaly", r"big data", r"data engineer", r"\bbi\b", r"analytics",
]), re.I)

SW_RX = re.compile("|".join([
    r"full ?-?stack", r"frontend", r"front-end", r"backend", r"back-end",
    r"webentwickl", r"web-entwickl", r"\breact\b", r"\bangular\b", r"\bvue\b",
    r"node\.?js", r"javascript", r"typescript", r"\bjava\b", r"\bphp\b",
    r"\.net\b", r"\bc#", r"softwareentwickl", r"software engineer",
    r"\bdevops\b", r"app-?entwickl", r"\bapi\b", r"\bcloud\b",
    r"\bentwickler", r"developer", r"\bgame\b", r"informatik",
    r"\bcoding\b", r"\bembedded\b",
]), re.I)

# --- Tabellenaufbau -------------------------------------------------------
COLUMNS = [
    ("Gefunden am", 12), ("Relevanz", 10), ("Titel", 50), ("Unternehmen", 26),
    ("Ort", 18), ("Entfernung", 11), ("Remote", 9), ("Art", 17),
    ("Dauer", 13), ("Eintritt", 12), ("Veröffentlicht", 13),
    ("Suchprofil", 26), ("CV-Empfehlung", 32), ("Link", 44),
    ("Referenz", 20), ("Status", 18), ("Notizen", 34), ("Bewerbung", 26),
    ("Punkte", 8),
]
COL = {name: i for i, (name, _) in enumerate(COLUMNS, start=1)}

STATI = ("Neu,Interessant,Beworben,Eingangsbestätigung,Gespräch,"
         "Zusage,Absage,Keine Rückmeldung,Verworfen")


def first(d, *keys, default=""):
    """Ersten vorhandenen, nicht leeren Wert aus d zurückgeben."""
    for k in keys:
        v = d.get(k)
        if v not in (None, "", [], {}):
            return v
    return default


def fetch(search):
    """Ein Suchprofil gegen die API laufen lassen."""
    params = {k: v for k, v in search.items()
              if k not in ("label", "remote", "pages")}
    params.setdefault("veroeffentlichtseit", MAX_AGE_DAYS)
    params["size"] = PAGE_SIZE
    # Remote-Profile suchen bundesweit und werfen ~95 % wieder weg,
    # also brauchen sie mehr Seiten, um überhaupt etwas zu finden.
    pages = search.get("pages") or (6 if search.get("remote") else MAX_PAGES)
    out = []
    for page in range(1, pages + 1):
        params["page"] = page
        try:
            r = requests.get(API, headers=HEADERS, params=params, timeout=30)
        except requests.RequestException as e:
            print(f"  ! Netzwerkfehler bei '{search['label']}': {e}")
            break
        if r.status_code != 200:
            print(f"  ! HTTP {r.status_code} bei '{search['label']}'"
                  f"{' — ' + r.text[:120] if r.text else ''}")
            break
        data = r.json()
        # v6: ergebnisliste (fehlt bei null Treffern). v4 hiess stellenangebote.
        hits = data.get("ergebnisliste") or data.get("stellenangebote") or []
        out.extend(hits)
        if len(hits) < PAGE_SIZE:
            break
        time.sleep(0.4)
    return out


def is_remote(hit):
    v = first(hit, "homeofficemoeglich", "homeOfficeMoeglich", default=None)
    if isinstance(v, bool):
        return v
    return str(v).strip().lower() in ("true", "ja", "1", "y")


def recommend_cv(titel):
    """Welchen der beiden Lebensläufe anhängen.

    Nur der Stellentitel zählt. `hauptberuf` ist die Einordnung des
    Arbeitgebers und oft irreführend — die Vodafone-Praktikumsstelle für
    Mobile Media läuft dort als 'Finanzassistent/in'.
    """
    ki = len(KI_RX.findall(titel))
    sw = len(SW_RX.findall(titel))
    if ki > sw:
        return f"KI — {CV_KI}"
    if sw > ki:
        return f"Software — {CV_SW}"
    return f"Software (unklar) — {CV_SW}"


def _von(hit, key):
    """Zeitraum-Felder in v6 sind {'von': 'YYYY-MM-DD'}."""
    z = hit.get(key)
    if isinstance(z, dict):
        return str(z.get("von") or "")[:10]
    return str(z or "")[:10]


def dauer(hit):
    """Vertragsdauer in einem Wort."""
    n = hit.get("befristungInMonaten")
    if n:
        return f"{n} Monate"
    v = (hit.get("vertragsdauer") or "").upper()
    if v == "UNBEFRISTET":
        return "unbefristet"
    if v == "BEFRISTET":
        return "befristet"
    return ""


def _alter_tage(datum):
    """Wie viele Tage ist die Veröffentlichung her? None, wenn kein Datum."""
    try:
        return (dt.date.today() - dt.date.fromisoformat(str(datum)[:10])).days
    except (ValueError, TypeError):
        return None


def punkte(row):
    """Wie gut passt die Stelle zum Lebenslauf? Höher ist besser.

    Grundlage sind die Fachbegriffe im Titel. Nähe, Vertragsdauer und
    Frische kommen als Aufschlag dazu — aber nur, wenn überhaupt ein
    fachlicher Bezug da ist. Sonst käme eine Marketingstelle um die Ecke
    allein über die Entfernung nach oben.
    """
    titel = row.get("titel") or ""
    basis = sum(n for rx, n in TECH_PUNKTE if rx.search(titel))
    if basis == 0:
        return 0
    basis = min(basis, TECH_DECKEL)

    bonus = 0

    # Nähe — Wohnort ist Krefeld
    km = row.get("entfernung_km")
    if row.get("remote"):
        bonus += 8
    elif km is not None:
        bonus += 10 if km <= 25 else 6 if km <= 50 else 3 if km <= 80 else 0

    # Praxisphase: Pflichtpraktika und lange Verträge bringen Stunden
    if PFLICHT_RX.search(titel):
        bonus += 8
    m = re.match(r"(\d+)\s*Monate", str(row.get("dauer") or ""))
    if m:
        n = int(m.group(1))
        bonus += 6 if n >= 6 else 3 if n >= 3 else 0
    elif row.get("dauer") == "unbefristet":
        bonus += 6

    # Frische
    alter = _alter_tage(row.get("veroeffentlicht"))
    if alter is not None:
        bonus += 6 if alter <= 7 else 3 if alter <= 14 else -4 if alter > 45 else 0

    # Die Zuschläge dürfen den fachlichen Kern nicht überstimmen. Sonst
    # steigt eine IT-Support-Stelle um die Ecke allein über Nähe und
    # Frische auf "hoch". Abzüge greifen dagegen ungedeckelt.
    return max(basis + min(bonus, basis), 1)


def relevanz(p):
    """Punktzahl in die drei Stufen der Tabellenspalte übersetzen."""
    if p >= SCHWELLE_HOCH:
        return "hoch"
    if p >= SCHWELLE_MITTEL:
        return "mittel"
    return "niedrig"


def normalise(hit, label):
    """Einen Treffer auf unsere Spalten herunterbrechen."""
    # v6-Feldnamen, mit v4 als Rückfall
    titel = first(hit, "stellenangebotsTitel", "titel", "beruf")
    ref = first(hit, "referenznummer", "refnr", "hashId")

    lok = first(hit, "stellenlokationen", default=[])
    adr = (lok[0].get("adresse") or {}) if isinstance(lok, list) and lok else {}
    alt_ort = first(hit, "arbeitsort", default={})
    ort = first(adr, "ort") or (alt_ort.get("ort", "") if isinstance(alt_ort, dict) else "")
    if ort.isupper():
        ort = ort.title()

    entf = hit.get("entfernung")
    if entf in (None, "") and isinstance(alt_ort, dict):
        entf = alt_ort.get("entfernung")

    veroeff = (_von(hit, "veroeffentlichungszeitraum")
               or str(first(hit, "datumErsteVeroeffentlichung",
                            "aktuelleVeroeffentlichungsdatum"))[:10])

    art_raw = (hit.get("stellenangebotsart") or "").upper()

    d = {
        "titel": titel or first(hit, "hauptberuf"),
        "firma": first(hit, "firma", "arbeitgeber"),
        "ort": ort,
        "entfernung": f"{entf} km" if entf not in (None, "") else "",
        "entfernung_km": int(entf) if str(entf).isdigit() else None,
        "remote": "ja" if is_remote(hit) else "",
        "art": ARTEN.get(art_raw, art_raw.replace("_", " ").title()),
        "art_raw": art_raw,
        "dauer": dauer(hit),
        "eintritt": (_von(hit, "eintrittszeitraum")
                     or str(first(hit, "eintrittsdatum", "eintrittstermin"))[:10]),
        "veroeffentlicht": veroeff,
        "profil": label,
        "cv": recommend_cv(titel),
        "link": DETAIL_URL.format(ref=ref) if ref else first(hit, "externeUrl"),
        "ref": ref,
    }
    d["punkte"] = punkte(d)
    d["relevanz"] = relevanz(d["punkte"])
    return d


def too_old(row, max_age):
    """v6 ignoriert veroeffentlichtseit, also prüfen wir das Datum selbst."""
    try:
        d = dt.date.fromisoformat(row["veroeffentlicht"])
    except (ValueError, TypeError):
        return False          # kein Datum: lieber zeigen als verwerfen
    return (dt.date.today() - d).days > max_age


def wanted(row):
    t = row["titel"] or ""
    if not row["ref"] or row.get("art_raw") in ARTEN_BLOCKLIST:
        return False
    if any(b in t.lower() for b in TITLE_BLOCKLIST):
        return False
    # Studierendenstelle? Entweder sagt der Titel es, oder die Angebotsart.
    if not STUDENT_RX.search(t) and row.get("art_raw") != "PRAKTIKUM_TRAINEE":
        return False
    return True


def open_or_create(path):
    """Tabelle laden oder mit formatierten Kopfzeilen neu anlegen."""
    try:
        wb = load_workbook(path)
        ws = wb["Funde"]
        # Fehlende Spalten am Ende ergänzen, statt die Datei zu verwerfen.
        thin = Side(style="thin", color="D8DDE2")
        ergaenzt = []
        for i, (name, width) in enumerate(COLUMNS, start=1):
            if ws.cell(row=1, column=i).value != name:
                if ws.cell(row=1, column=i).value:
                    print(f"  ! Spalte {i} heisst '{ws.cell(row=1, column=i).value}', "
                          f"erwartet '{name}'. Datei umbenennen und neu starten.")
                    sys.exit(1)
                c = ws.cell(row=1, column=i, value=name)
                c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="12333A")
                c.alignment = Alignment(vertical="center", wrap_text=True)
                c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                ws.column_dimensions[get_column_letter(i)].width = width
                ergaenzt.append(name)
        if ergaenzt:
            print(f"  Spalten ergänzt: {', '.join(ergaenzt)}")
            # Autofilter mitziehen, sonst endet er vor den neuen Spalten
            ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"
        return wb, ws, False
    except (FileNotFoundError, KeyError):
        pass

    wb = Workbook()
    ws = wb.active
    ws.title = "Funde"
    thin = Side(style="thin", color="D8DDE2")
    for i, (name, width) in enumerate(COLUMNS, start=1):
        c = ws.cell(row=1, column=i, value=name)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="12333A")
        c.alignment = Alignment(vertical="center", wrap_text=True)
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}1"

    dv = DataValidation(type="list", formula1=f'"{STATI}"', allow_blank=True)
    ws.add_data_validation(dv)
    letter = get_column_letter(COL["Status"])
    dv.add(f"{letter}2:{letter}3000")

    rel = get_column_letter(COL["Relevanz"])
    span = f"{rel}2:{rel}3000"
    for wort, hg, fg in (("hoch", "D6ECDF", "14603F"),
                         ("mittel", "FBF0D4", "7A5410"),
                         ("niedrig", "EFEFEF", "7A7A7A")):
        ws.conditional_formatting.add(span, CellIsRule(
            operator="equal", formula=[f'"{wort}"'],
            fill=PatternFill("solid", bgColor=hg),
            font=Font(name="Arial", size=10, bold=(wort == "hoch"), color=fg)))
    return wb, ws, True


LEGAL_RX = re.compile(r"\b(gmbh|ag|se|ug|kg|mbh|co|kgaa|e\.?v|holding|group|"
                      r"gruppe|deutschland|germany|und|&)\b", re.I)


def fkey(name):
    """Firmenname auf einen Vergleichsschlüssel reduzieren."""
    if not name:
        return ""
    return re.sub(r"[^a-z0-9]", "", LEGAL_RX.sub(" ", str(name).lower()))


# Wörter, die in fast jedem Stellentitel stehen und daher nichts unterscheiden.
FUELLWOERTER = {
    "praktikant", "praktikantin", "praktikanten", "praktikum", "praktika",
    "pflichtpraktikum", "werkstudent", "werkstudentin", "werkstudenten",
    "student", "studentin", "studentische", "studentischer", "working",
    "trainee", "intern", "internship", "mwd", "wmd", "mfd", "wmdgn", "gn",
    "genders", "all", "mit", "den", "dem", "der", "die", "das", "und", "im",
    "in", "bei", "für", "als", "bereich", "bereichen", "schwerpunkt",
    "schwerpunkten", "team", "abteilung", "sowie", "oder", "einer", "eines",
    "zur", "zum", "von", "auf", "the", "for", "and", "with", "of",
}


def ptokens(pos):
    """Stellentitel auf seine Inhaltswörter reduzieren."""
    if not pos:
        return frozenset()
    s = re.sub(r"\(.*?\)", " ", str(pos).lower())
    s = re.sub(r"m/w/d|w/m/d|m/f/d|all genders", " ", s)
    woerter = re.split(r"[^a-zA-ZäöüßÄÖÜ0-9]+", s)
    return frozenset(w for w in woerter
                     if len(w) >= 2 and w not in FUELLWOERTER)


def dieselbe_stelle(a, b):
    """Bezeichnen zwei Titel dieselbe Stelle?

    Ein Zeichenvergleich trennt das nicht: 'Innovation Garage IoT Backend
    Development' und 'Innovation Garage Cyber Security' liegen bei 0,82,
    derselbe IoT-Titel anders formuliert bei 0,88 — drei Hundertstel
    Abstand. Also über die Inhaltswörter: geteilt durch die kleinere
    Wortmenge, damit ein verkürzter Titel ('Quality Engineering' statt
    'Quality Engineering & Process Automation') noch zuordenbar bleibt.
    """
    if not a or not b:
        return False
    if a == b:
        return True
    gemeinsam = len(a & b)
    kleiner = min(len(a), len(b))
    if kleiner < 2:                      # zu wenig Information
        return gemeinsam == kleiner and gemeinsam > 0
    return gemeinsam / kleiner >= 0.70


def load_applied(path):
    """Arbeitgeber und Positionen aus Bewerbungen.xlsx lesen.

    Liefert eine Liste von (firmenschlüssel, positionsschlüssel, Beschriftung).
    Ohne die Datei kommt eine leere Liste zurück — das Skript läuft dann
    einfach ohne Dublettenprüfung weiter.
    """
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except (FileNotFoundError, KeyError, OSError):
        return []
    out = []
    for name in wb.sheetnames:
        if not re.fullmatch(r"\d{4}-\d{2}", name):
            continue
        ws = wb[name]
        for row in ws.iter_rows(min_row=4, max_col=9, values_only=True):
            firma, pos, status = row[2], row[3], row[8]
            if not firma and not pos:
                continue
            if str(status or "").strip() in ("Verworfen",):
                continue
            out.append((fkey(firma), ptokens(pos),
                        f"{firma or '?'} ({name[5:]}.{name[:4]}, {status or '—'})"))
    wb.close()
    return out


def applied_state(row, applied):
    """'' | ('andere Stelle', Text) | ('dieselbe Stelle', Text)"""
    fk, pk = fkey(row["firma"]), ptokens(row["titel"])
    if not fk:
        return None
    gleiche_firma = [a for a in applied
                     if a[0] and (a[0] in fk or fk in a[0])]
    if not gleiche_firma:
        return None
    for a in gleiche_firma:
        if dieselbe_stelle(a[1], pk):
            return ("dieselbe Stelle", a[2])
    return ("andere Stelle", gleiche_firma[0][2])


def known_refs(ws):
    c = COL["Referenz"]
    return {ws.cell(row=r, column=c).value
            for r in range(2, ws.max_row + 1)
            if ws.cell(row=r, column=c).value}


def zeile_schreiben(ws, r, werte):
    """Eine Datenzeile setzen und formatieren."""
    thin = Side(style="thin", color="E7ECEF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    umbruch = (COL["Titel"], COL["CV-Empfehlung"], COL["Notizen"],
               COL["Bewerbung"])
    for i in range(1, len(COLUMNS) + 1):
        c = ws.cell(row=r, column=i,
                    value=werte[i - 1] if i <= len(werte) else None)
        c.font = Font(name="Arial", size=10)
        c.alignment = Alignment(vertical="top", wrap_text=(i in umbruch))
        c.border = border
    ws.cell(row=r, column=COL["Gefunden am"]).number_format = "DD.MM.YYYY"
    link = ws.cell(row=r, column=COL["Link"])
    if link.value:
        link.hyperlink = str(link.value)
        link.font = Font(name="Arial", size=10, color="0D6B6E",
                         underline="single")
    if ws.cell(row=r, column=COL["Bewerbung"]).value:
        ws.cell(row=r, column=COL["Bewerbung"]).font = Font(
            name="Arial", size=10, color="8C2F2F")
    ws.cell(row=r, column=COL["Punkte"]).alignment = Alignment(
        vertical="top", horizontal="center")
    ws.row_dimensions[r].height = 30


def km_aus_text(v):
    m = re.match(r"\s*(\d+)", str(v or ""))
    return int(m.group(1)) if m else None


def neu_bewerten(ws):
    """Punkte und Relevanz für alle Zeilen neu berechnen.

    Das muss bei jedem Lauf passieren und nicht nur für neue Funde: Der
    Frische-Aufschlag verfällt mit den Tagen. Eine im Oktober gefundene
    Stelle soll heute nicht mehr so weit oben stehen wie damals.
    """
    for r in range(2, ws.max_row + 1):
        titel = ws.cell(row=r, column=COL["Titel"]).value
        if not titel:
            continue
        p = punkte({
            "titel": titel,
            "remote": ws.cell(row=r, column=COL["Remote"]).value,
            "entfernung_km": km_aus_text(
                ws.cell(row=r, column=COL["Entfernung"]).value),
            "dauer": ws.cell(row=r, column=COL["Dauer"]).value,
            "veroeffentlicht": ws.cell(row=r, column=COL["Veröffentlicht"]).value,
        })
        ws.cell(row=r, column=COL["Punkte"]).value = p
        ws.cell(row=r, column=COL["Relevanz"]).value = relevanz(p)


def sortiere(ws, modus):
    """Alle Datenzeilen neu ordnen. Zeile 1 bleibt die Kopfzeile."""
    daten = []
    for r in range(2, ws.max_row + 1):
        werte = [ws.cell(row=r, column=c).value
                 for c in range(1, len(COLUMNS) + 1)]
        if any(v not in (None, "") for v in werte):
            daten.append(werte)
    if len(daten) < 2:
        return 0

    i_pkt, i_ver = COL["Punkte"] - 1, COL["Veröffentlicht"] - 1

    def tag(z):
        v = z[i_ver]
        if isinstance(v, dt.datetime):      # vor dt.date prüfen, Unterklasse
            return v.date().toordinal()
        if isinstance(v, dt.date):
            return v.toordinal()
        try:
            return dt.date.fromisoformat(str(v)[:10]).toordinal()
        except (ValueError, TypeError):
            return 0

    def pkt(z):
        return z[i_pkt] if isinstance(z[i_pkt], (int, float)) else 0

    if modus == "datum":
        daten.sort(key=lambda z: (-tag(z), -pkt(z)))
    else:
        daten.sort(key=lambda z: (-pkt(z), -tag(z)))

    for n, werte in enumerate(daten, start=2):
        zeile_schreiben(ws, n, werte)
    for r in range(len(daten) + 2, ws.max_row + 1):     # Reste leeren
        for c in range(1, len(COLUMNS) + 1):
            ws.cell(row=r, column=c).value = None
    return len(daten)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="Jobfunde.xlsx")
    ap.add_argument("--sortierung", choices=("relevanz", "datum"),
                    default="relevanz",
                    help="Reihenfolge in der Tabelle: 'relevanz' sortiert nach "
                         "Punkten und bei Gleichstand nach Datum, 'datum' "
                         "umgekehrt (Standard: relevanz)")
    ap.add_argument("--max-age", type=int, default=MAX_AGE_DAYS,
                    help=f"Anzeigen älter als N Tage verwerfen (Standard {MAX_AGE_DAYS})")
    ap.add_argument("--bewerbungen",
                    default=str(Path.home() / "Jobsuche" / "Bewerbungen.xlsx"),
                    help="Bewerbungstabelle für die Dublettenprüfung")
    ap.add_argument("--debug", action="store_true",
                    help="Rohfelder des ersten Treffers zeigen und beenden")
    args = ap.parse_args()

    applied = load_applied(args.bewerbungen)
    if applied:
        firmen = len({a[0] for a in applied if a[0]})
        print(f"Bewerbungen: {len(applied)} Vorgänge bei {firmen} Arbeitgebern "
              f"aus {Path(args.bewerbungen).name}")
    elif Path(args.bewerbungen).exists():
        print(f"{Path(args.bewerbungen).name} enthält noch keine Monatsblätter "
              f"(Blattname JJJJ-MM) — Lauf ohne Dublettenprüfung.")
    else:
        print(f"Keine Bewerbungstabelle unter {args.bewerbungen} — "
              f"Lauf ohne Dublettenprüfung.")

    wb, ws, created = open_or_create(args.out)
    seen = known_refs(ws)
    print(f"Tabelle: {args.out} "
          f"({'neu angelegt' if created else f'{len(seen)} bekannte Einträge'})")

    found, skip_remote, skip_old, skip_far, skip_dup, new = 0, 0, 0, 0, 0, []
    for s in SEARCHES:
        hits = fetch(s)
        found += len(hits)

        if args.debug and hits:
            print("\n--- Rohfelder des ersten Treffers ---")
            for k, v in hits[0].items():
                print(f"  {k}: {v}")
            print("\n--- so gelesen ---")
            for k, v in normalise(hits[0], s["label"]).items():
                print(f"  {k}: {v}")
            return

        kept = 0
        for h in hits:
            if s.get("remote") and not is_remote(h):
                skip_remote += 1
                continue
            row = normalise(h, s["label"])
            if row["ref"] in seen or not wanted(row):
                continue
            if too_old(row, args.max_age):
                skip_old += 1
                continue
            # umkreis wird von v6 offenbar nicht zuverlässig beachtet
            # (Stuttgart tauchte in einer Suche um Krefeld auf), also
            # prüfen wir die Entfernung selbst, wenn die API sie mitliefert.
            grenze = s.get("umkreis")
            km = row["entfernung_km"]
            if grenze and km is not None and km > grenze * 1.25 \
                    and not row["remote"]:
                skip_far += 1
                continue
            seen.add(row["ref"])
            new.append(row)
            kept += 1
        print(f"  {s['label']}: {len(hits)} Treffer, {kept} neu")

    today = dt.date.today()
    r = ws.max_row + 1
    for row in new:
        zeile_schreiben(ws, r, [
            today, row["relevanz"], row["titel"], row["firma"],
            row["ort"], row["entfernung"], row["remote"], row["art"],
            row["dauer"], row["eintritt"], row["veroeffentlicht"],
            row["profil"], row["cv"], row["link"], row["ref"],
            "Neu", "", "", row["punkte"],
        ])
        r += 1

    # Spalte "Bewerbung" für ALLE Zeilen neu bestimmen, nicht nur für neue.
    # Der Bewerbungsstand ändert sich täglich; ein beim Fund einmal
    # geschriebener Hinweis wäre nach einer Woche falsch.
    gepflegt = 0
    if applied:
        for r in range(2, ws.max_row + 1):
            firma = ws.cell(row=r, column=COL["Unternehmen"]).value
            titel = ws.cell(row=r, column=COL["Titel"]).value
            if not firma and not titel:
                continue
            z = applied_state({"firma": firma, "titel": titel}, applied)
            if not z:
                text = ""
            elif z[0] == "dieselbe Stelle":
                text = f"beworben: {z[1]}"
            else:
                text = f"dort schon beworben: {z[1]}"
            zelle = ws.cell(row=r, column=COL["Bewerbung"])
            if (zelle.value or "") != text:      # openpyxl legt "" als None ab
                zelle.value = text
                zelle.font = Font(name="Arial", size=10,
                                  color="8C2F2F" if text else "26313C")
                zelle.alignment = Alignment(vertical="top", wrap_text=True)
                gepflegt += 1

    neu_bewerten(ws)
    geordnet = sortiere(ws, args.sortierung)

    wb.save(args.out)
    verworfen = []
    if skip_remote:
        verworfen.append(f"{skip_remote} ohne Homeoffice")
    if skip_old:
        verworfen.append(f"{skip_old} älter als {args.max_age} Tage")
    if skip_far:
        verworfen.append(f"{skip_far} außerhalb des Umkreises")
    if skip_dup:
        verworfen.append(f"{skip_dup} bereits beworben")
    if gepflegt:
        print(f"\nSpalte 'Bewerbung' bei {gepflegt} Zeilen aktualisiert.")
    if geordnet:
        print(f"{geordnet} Zeilen neu bewertet und nach "
              f"{'Datum' if args.sortierung == 'datum' else 'Passung'} sortiert.")
    print(f"\n{found} Treffer geprüft, {len(new)} neue Stellen ergänzt"
          f"{' (' + ', '.join(verworfen) + ' verworfen)' if verworfen else ''}.")
    if new:
        def rang(z):
            t = _alter_tage(z["veroeffentlicht"])
            t = -t if t is not None else -9999
            return (t, z["punkte"]) if args.sortierung == "datum" \
                else (z["punkte"], t)

        beste = sorted(new, key=rang, reverse=True)
        wonach = "Datum" if args.sortierung == "datum" else "Passung"
        print(f"\nNeue Funde, nach {wonach} sortiert:")
        print(f"  {'Pkt':>3}  {'Titel':46} | {'Unternehmen':22} | "
              f"{'Ort':12} | Veröffentlicht")
        for row in beste[:20]:
            print(f"  {row['punkte']:>3}  {(row['titel'] or '—')[:46]:46} | "
                  f"{(row['firma'] or '—')[:22]:22} | "
                  f"{(row['ort'] or '—')[:12]:12} | {row['veroeffentlicht']}")
        if len(beste) > 20:
            print(f"  … und {len(beste) - 20} weitere in der Datei.")


if __name__ == "__main__":
    sys.exit(main())
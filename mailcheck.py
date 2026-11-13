#!/usr/bin/env python3
"""
mailcheck.py — liest täglich das Gmail-Postfach und schreibt die
Bewerbungstabelle fort.

Was es tut
----------
* sucht im Posteingang nach Eingangsbestätigungen, Absagen, Einladungen
  und Erinnerungen an unvollständige Bewerbungen
* sucht im Gesendet-Ordner nach Bewerbungen, die du direkt per Mail
  verschickt hast
* erkennt Arbeitgeber und Position aus dem Text
* findet die passende Zeile in Bewerbungen.xlsx und schreibt Status,
  letzte Aktion und eine Notiz fort
* legt für einen neuen Monat selbst ein neues Blatt an
* verschickt nichts und antwortet auf nichts

Einrichten
----------
1. Zwei-Faktor-Anmeldung im Google-Konto aktivieren.
2. Unter myaccount.google.com/apppasswords ein App-Passwort erzeugen.
3. Zugangsdaten ablegen:

       cat > ~/.jobsuche-mail.json <<'EOF'
       {"user": "deine.adresse@gmail.com", "password": "APP-PASSWORT"}
       EOF
       chmod 600 ~/.jobsuche-mail.json

   Alternativ die Umgebungsvariablen GMAIL_USER und GMAIL_APP_PASSWORD.

Benutzung
---------
    python3 mailcheck.py --dry-run        # zeigt nur, was es täte
    python3 mailcheck.py                  # schreibt in die Tabelle
    python3 mailcheck.py --since 2026-10-01
"""

import argparse
import datetime as dt
import email
import email.header
import imaplib
import json
import os
import re
import sys
from email.utils import parsedate_to_datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HOME = Path.home()
XLSX = HOME / "Jobsuche" / "Bewerbungen.xlsx"
STATE = HOME / "Jobsuche" / ".mailcheck-state.json"
CRED = HOME / ".jobsuche-mail.json"

HEADER_ROW = 3          # Zeile mit den Spaltennamen
FIRST_DATA_ROW = 4

# Spaltennummern in Bewerbungen.xlsx
C_NR, C_DATUM, C_FIRMA, C_POS, C_ART, C_ORT, C_QUELLE, C_KONTAKT = 1, 2, 3, 4, 5, 6, 7, 8
C_STATUS, C_AKTION, C_SCHRITT, C_WV, C_NOTIZ = 9, 10, 11, 12, 13
N_COLS = 13

STATI = ("Beworben,Eingangsbestätigung,Gespräch,Gespräch geführt,"
         "Coding-Challenge,Zusage,Absage,Keine Rückmeldung,"
         "Entwurf unvollständig,Vorbereitet,Verworfen")

# Spaltennamen und -breiten. Nur gebraucht, wenn die Datei fehlt und ganz
# neu angelegt werden muss; sonst wird das vorhandene Blatt als Muster genommen.
SPALTEN = [("Nr.", 5), ("Beworben am", 12), ("Unternehmen", 27), ("Position", 44),
           ("Art", 16), ("Standort", 16), ("Quelle", 19), ("Kontaktperson", 18),
           ("Status", 21), ("Letzte Aktion", 13), ("Nächster Schritt", 28),
           ("Wiedervorlage", 13), ("Notizen", 42)]
AMPEL = (("Absage", "FBE3E3", "8C2F2F"), ("Zusage", "D6ECDF", "14603F"),
         ("Gespräch", "D6E4F5", "1B4A7A"), ("Coding-Challenge", "D6E4F5", "1B4A7A"),
         ("Entwurf unvollständig", "FBF0D4", "7A5410"))

# --- Absender, die nur Stellenwerbung schicken ----------------------------
IGNORE_SENDERS = [
    "jobagent.stepstone.de", "match.indeed.com", "jobalert.indeed.com",
    "jobalerts-noreply@linkedin.com", "quora.com", "substack.com",
    "jobs2web.com", "nl.", "newsletter", "noreply@medium.com",
    "jobs-noreply@linkedin.com_DIGEST",
]
IGNORE_SUBJECTS = [
    "looking for candidates like you", "new job opportunities",
    "our recommendation", "profile stands out", "jobs you may be interested",
    "is hiring for", "job at ", "may be a fit for", "closing soon",
    "chance for an interview for this job", "job-e-mail", "bestätigen sie ihre",
]

# --- Klassifizierung. Reihenfolge entscheidet. ----------------------------
RULES = [
    ("Absage", [
        r"leider nicht erfolgreich", r"leider nicht berücksichtig",
        r"nicht weiter berücksichtig", r"für andere bewerber entschieden",
        r"haben wir uns .{0,40}entschieden", r"do not have good news",
        r"unfortunately", r"not (?:been )?(?:selected|successful)",
        r"\babsage\b", r"eine andere entscheidung",
    ]),
    ("Gespräch", [
        r"einladung zu[rm]", r"vorstellungsgespräch", r"zum kennenlernen",
        r"kennenlernen!?", r"interview", r"terminbestätigung",
        r"gespräch", r"coding[- ]challenge", r"online[- ]assessment",
    ]),
    ("Entwurf unvollständig", [
        r"vervollständige", r"ergänze nur noch", r"bewerbung abschließen",
        r"fast geschafft", r"fast fertig", r"schließe jetzt deine bewerbung",
        r"es fehlen noch", r"lass dir den job nicht entgehen",
    ]),
    ("Eingangsbestätigung", [
        r"bewerbung.{0,30}(?:eingegangen|angekommen|erhalten)",
        r"(?:deine|ihre) bewerbung (?:für|auf) die stelle",
        r"dass (?:du dich|sie sich) für (?:einen|eine) (?:job|stelle|tätigkeit)",
        r"bedanken uns für (?:ihre|deine) bewerbung",
        r"(?:vielen )?dank für (?:ihre|deine) bewerbung",
        r"thanks? for applying", r"thanks for your interest",
        r"application .{0,90}has been sent", r"application was sent to",
        r"sie haben sich auf die ausschreibung", r"beworben hast",
        r"bewerbung über indeed", r"ist im kasten",
        r"freuen uns über (?:ihr|dein) interesse", r"we received your application",
    ]),
    ("Beworben", [
        r"noch in bearbeitung", r"nicht vergessen", r"wartezeit",
    ]),
]

# --- Arbeitgeber aus dem Text ziehen --------------------------------------
FIRMA_RX = [
    # "… bei beispiel-gmbh.de ist erfolgreich eingegangen" — auch klein geschrieben
    r"bei (?:der |dem )?([A-Za-zÄÖÜäöü][\w&.\'\-]*(?: [A-ZÄÖÜa-zäöüß&.\'\-]+){0,4}?) ist (?:erfolgreich eingegangen|bei uns angekommen)",
    r"[Bb]ewerbung bei (?:der |dem |die |das )?([A-Za-zÄÖÜäöü][\w&.\'\-]*(?: [A-ZÄÖÜa-zäöüß&.\'/\-]+){0,5}?)"
    r"(?= auf die Stelle| ist (?:noch )?in Bearbeitung| war | ist eingegangen|[,.!])",
    # Stoppwörter, sonst verschluckt das Muster " erhalten"
    r"Position bei (?:der )?([A-Za-zÄÖÜäöü][\w&.\'\-]*(?: [A-ZÄÖÜa-zäöüß&.\'\-]+){0,4}?)"
    r"(?= erhalten| eingegangen| ist\b| angekommen|[,.!]|$)",
    r"application was sent to ([A-Z][\w&.\'\-]*(?: [A-Za-z&.\'\-]+){0,4})",
    r"[Aa]usschreibung .{3,80}? bei ([A-ZÄÖÜ][\w&.\'\-]*(?: [A-ZÄÖÜa-zäöüß&.\'\-]+){0,3}?) beworben",
    r"interest in ([A-Z][\w&.\'\-]*(?: [A-Za-z&.\'\-]+){0,3})!",
    r"Teil (?:unseres|des) ([A-ZÄÖÜ][\w\-]+)-Teams",
]

POS_RX = [
    r"auf den Job (.{5,90}?)(?: bei | ab\b|\s*\n|$)",
    # Doppelpunkt ist Pflicht, sonst greift es in der Betreffzeile daneben
    r"für die Position:\s*(.{5,90}?)\s*(?:\n|$|Derzeit|Wir )",
    r"[Aa]pplication as (.{5,90}?)(?: I | at |\s*\n|$)",
    r"[Bb]ewerbung als (.{5,90}?)(?: bei | ist |\s*\n|$)",
    r"auf die Stelle [„\"'»]?(.{5,90}?)[\"'«”]?(?: geprüft| war |\s*\n|$)",
    r"[Aa]usschreibung (.{5,90}?) bei ",
    r"applying for the role of (.{5,90}?)(?:[.\n]|$)",
    r"application for (.{5,90}?) has been sent",
    r"[Bb]ewerbung über Indeed:\s*(.{5,90}?)(?:\n|$)",
    r"[Bb]ewerbung auf die Stelle (.{5,90}?)(?:\n|$)",
]

LEGAL = re.compile(r"\b(gmbh|ag|se|ug|kg|mbh|co|kgaa|e\.?v|holding|group|"
                   r"gruppe|deutschland|germany|und|&)\b", re.I)

ATS_DOMAINS = {"workwise.io", "join.com", "linkedin.com", "successfactors.eu",
               "indeed.com", "stepstone.de", "personio.de", "myworkday.com",
               "softgarden.io", "guidecom.de", "hrworks.de", "jobs2web.com",
               "smartrecruiters.com", "greenhouse.io", "recruitee.com",
               "concludis.de", "rexx-systems.com", "gmail.com", "google.com"}


# =========================================================================
# Hilfsfunktionen
# =========================================================================
def creds():
    """Zugangsdaten holen. Leerzeichen im App-Passwort werden entfernt —
    Google zeigt es in vier Blöcken, gemeint ist die Folge ohne Lücken."""
    u = os.environ.get("GMAIL_USER")
    p = os.environ.get("GMAIL_APP_PASSWORD")
    if not (u and p):
        if not CRED.exists():
            sys.exit(f"Keine Zugangsdaten. Lege {CRED} an "
                     f"(siehe Kopf dieser Datei).")
        d = json.loads(CRED.read_text())
        u, p = d.get("user", ""), d.get("password", "")
    p = re.sub(r"\s+", "", p or "")
    if not u or not p:
        sys.exit("Benutzername oder Passwort fehlt in den Zugangsdaten.")
    if len(p) != 16:
        print(f"  Hinweis: Das Passwort hat {len(p)} Zeichen. Ein Google-"
              f"App-Passwort hat genau 16. Wahrscheinlich steht dort noch "
              f"das normale Kontopasswort.")
    return u, p


def load_state():
    if STATE.exists():
        try:
            return json.loads(STATE.read_text())
        except json.JSONDecodeError:
            pass
    return {"last_run": None, "seen": []}


def decode(s):
    if not s:
        return ""
    out = []
    for part, enc in email.header.decode_header(s):
        if isinstance(part, bytes):
            out.append(part.decode(enc or "utf-8", errors="replace"))
        else:
            out.append(part)
    return " ".join(out).replace("\r", " ").replace("\n", " ").strip()


def body_text(msg):
    """Reinen Text aus der Mail holen, HTML notfalls grob entschlacken."""
    text = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            if ct == "text/plain":
                try:
                    text += part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8", errors="replace")
                except (AttributeError, LookupError):
                    pass
        if not text.strip():
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    try:
                        text += part.get_payload(decode=True).decode(
                            part.get_content_charset() or "utf-8", errors="replace")
                    except (AttributeError, LookupError):
                        pass
    else:
        try:
            text = msg.get_payload(decode=True).decode(
                msg.get_content_charset() or "utf-8", errors="replace")
        except (AttributeError, LookupError):
            text = str(msg.get_payload())
    text = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", text,
                  flags=re.S | re.I)
    text = re.sub(r"<[^>]+>", " ", text)
    text = (text.replace("&nbsp;", " ").replace("&amp;", "&")
                .replace("&quot;", '"').replace("&#39;", "'")
                .replace("&uuml;", "ü").replace("&auml;", "ä")
                .replace("&ouml;", "ö").replace("&szlig;", "ß"))
    return re.sub(r"[ \t ͏‌‍]{2,}", " ", text)[:6000]


def noise(sender, subject):
    s, sub = sender.lower(), subject.lower()
    if any(x in s for x in IGNORE_SENDERS):
        return True
    return any(x in sub for x in IGNORE_SUBJECTS)


def classify(subject, body):
    hay = f"{subject}\n{body}".lower()
    for status, patterns in RULES:
        for p in patterns:
            if re.search(p, hay, re.I):
                return status, p
    return None, None


def firma_aus_text(subject, body, sender):
    for rx in FIRMA_RX:
        for hay in (subject, body):
            m = re.search(rx, hay)
            if m:
                name = re.sub(r"\s+", " ", m.group(1)).strip(" .,:;!-")
                # Der Punkt steht in der Zeichenklasse, damit "Co." passt —
                # dadurch kann das Muster über einen Satzpunkt hinauslaufen.
                # Am Satzende abschneiden, aber Abkürzungen verschonen:
                # "Co.", "u.", "Dr." sind kurz, ein Satzende steht nach
                # einem längeren Wort.
                name = re.split(r"(?<=\w{5})\.\s+(?=[A-ZÄÖÜ])", name)[0]
                name = name.strip(" .,:;!-")
                erstes = name.split(" ")[0].lower().strip(".,;:!") if name else ""
                if erstes in ("uns", "euch", "dir", "ihnen", "us", "unserem",
                              "unserer", "unseren", "unser", "deiner", "deinem"):
                    continue
                if 2 < len(name) < 46:
                    return name
    # Rückfall: die Absenderadresse auswerten
    m = re.search(r"([\w.\-+]+)@([\w.\-]+)", sender)
    if m:
        local, host = m.group(1).lower(), m.group(2).lower()
        parts = host.split(".")
        ist_portal = any(".".join(parts[cut:]) in ATS_DOMAINS
                         for cut in range(len(parts) - 1))
        if ist_portal:
            # Portale wie Personio setzen den Arbeitgeber in den lokalen Teil:
            # adi-solutions-gmbh-jobs@m.personio.de -> ADI Solutions GmbH
            rest = re.sub(r"[-_.](jobs?|bewerbung\w*|karriere|recruiting|"
                          r"hr|noreply|no-reply|mail)$", "", local)
            woerter = [w for w in re.split(r"[-_.+]", rest) if len(w) > 1]
            if len(woerter) >= 2:
                FORM = {"gmbh": "GmbH", "ag": "AG", "se": "SE",
                        "ug": "UG", "kg": "KG", "mbh": "mbH"}
                name = " ".join(FORM.get(w, w.title()) for w in woerter)
                return name
            return None
        label = parts[-2] if len(parts) >= 2 else parts[0]
        label = re.sub(r"^(mail|email|noreply|no-reply|bewerbung\w*|karriere|"
                       r"jobs?|hr|recruiting|m|ds)$", "", label)
        if len(label) > 2:
            return label.replace("-", " ").title()
    return None


def position_aus_text(subject, body):
    for rx in POS_RX:
        for hay in (subject, body):
            m = re.search(rx, hay)
            if m:
                p = re.sub(r"\s+", " ", m.group(1)).strip(" .,:;!-\"'„“»«")
                p = re.sub(r"\s*\((?:m/w/d|w/m/d|m/f/d|all genders|gn)\)\s*$", "",
                           p, flags=re.I)
                if 4 < len(p) < 95:
                    return p
    return None


def key(name):
    """Firmenname auf einen Vergleichsschlüssel reduzieren."""
    if not name:
        return ""
    s = LEGAL.sub(" ", name.lower())
    return re.sub(r"[^a-z0-9]", "", s)


def poskey(p):
    if not p:
        return ""
    return re.sub(r"[^a-z0-9]", "", p.lower())[:40]


# =========================================================================
# IMAP
# =========================================================================
def sent_folder(imap):
    """Den Gesendet-Ordner finden, egal wie er in der Kontosprache heisst."""
    ok, data = imap.list()
    if ok != "OK":
        return None
    for raw in data:
        line = raw.decode(errors="replace")
        if "\\Sent" in line:
            m = re.search(r'"([^"]+)"\s*$', line)
            if m:
                return m.group(1)
    # Deutsche Konten heissen "[Gmail]/Gesendet", englische "Sent Mail"
    for kandidat in ("[Gmail]/Gesendet", "[Gmail]/Sent Mail"):
        if imap.select(f'"{kandidat}"', readonly=True)[0] == "OK":
            return kandidat
    return None


def fetch_folder(imap, folder, since):
    """Alle Mails eines Ordners ab einem Datum holen."""
    ok, _ = imap.select(f'"{folder}"', readonly=True)
    if ok != "OK":
        print(f"  ! Ordner '{folder}' nicht lesbar")
        return []
    crit = since.strftime("%d-%b-%Y")
    ok, data = imap.search(None, "SINCE", crit)
    if ok != "OK" or not data or not data[0]:
        return []
    ids = data[0].split()
    out = []
    for i in range(0, len(ids), 40):          # in Blöcken holen
        block = b",".join(ids[i:i + 40])
        ok, raw = imap.fetch(block, "(RFC822)")
        if ok != "OK":
            continue
        for item in raw:
            if isinstance(item, tuple) and item[1]:
                try:
                    out.append(email.message_from_bytes(item[1]))
                except Exception:
                    pass
    return out


# =========================================================================
# Tabelle
# =========================================================================
def month_sheets(wb):
    return sorted(n for n in wb.sheetnames if re.fullmatch(r"\d{4}-\d{2}", n))


def ensure_month(wb, ym):
    """Blatt für JJJJ-MM liefern, bei Bedarf nach dem Muster des letzten anlegen."""
    if ym in wb.sheetnames:
        return wb[ym], False
    vorhanden = month_sheets(wb)
    muster = wb[vorhanden[-1]] if vorhanden else None
    ws = wb.create_sheet(ym)
    # Monatsblätter stehen vorn und chronologisch; neues Blatt einsortieren
    ziel = month_sheets(wb).index(ym)
    jetzt = wb.sheetnames.index(ym)
    if jetzt != ziel:
        wb._sheets.insert(ziel, wb._sheets.pop(jetzt))

    monat = dt.date(int(ym[:4]), int(ym[5:]), 1)
    namen = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
             "August", "September", "Oktober", "November", "Dezember"]
    ws.cell(row=1, column=1,
            value=f"Bewerbungen · {namen[monat.month - 1]} {monat.year}").font = \
        Font(name="Arial", size=13, bold=True, color="12333A")
    ws.cell(row=2, column=1,
            value=f"Von mailcheck.py angelegt am "
                  f"{dt.date.today().strftime('%d.%m.%Y')}.").font = \
        Font(name="Arial", size=9, italic=True, color="6B7A85")

    thin = Side(style="thin", color="D8DDE2")
    for c in range(1, N_COLS + 1):
        name = (muster.cell(row=HEADER_ROW, column=c).value if muster
                else SPALTEN[c - 1][0])
        cell = ws.cell(row=HEADER_ROW, column=c, value=name)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="12333A")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        letter = get_column_letter(c)
        ws.column_dimensions[letter].width = (
            muster.column_dimensions[letter].width if muster
            else SPALTEN[c - 1][1])
    ws.row_dimensions[HEADER_ROW].height = 26
    ws.freeze_panes = f"C{FIRST_DATA_ROW}"
    ws.auto_filter.ref = f"A{HEADER_ROW}:{get_column_letter(N_COLS)}{HEADER_ROW}"

    dv = DataValidation(type="list", formula1=f'"{STATI}"', allow_blank=True)
    ws.add_data_validation(dv)
    span = f"I{FIRST_DATA_ROW}:I600"
    dv.add(span)
    for wort, bg, fg in AMPEL:
        ws.conditional_formatting.add(span, CellIsRule(
            operator="equal", formula=[f'"{wort}"'],
            fill=PatternFill("solid", bgColor=bg),
            font=Font(name="Arial", size=10, bold=True, color=fg)))
    return ws, True


def read_rows(wb):
    """Alle Datenzeilen aller Monatsblätter einsammeln."""
    rows = []
    for name in month_sheets(wb):
        ws = wb[name]
        for r in range(FIRST_DATA_ROW, ws.max_row + 1):
            firma = ws.cell(row=r, column=C_FIRMA).value
            pos = ws.cell(row=r, column=C_POS).value
            if not firma and not pos:
                continue
            rows.append({"sheet": name, "row": r, "firma": firma, "pos": pos,
                         "status": ws.cell(row=r, column=C_STATUS).value})
    return rows


def find_row(rows, firma, pos):
    """Zeile zu Arbeitgeber und Position finden. Firma zählt mehr als Position."""
    fk, pk = key(firma), poskey(pos)
    if fk:
        treffer = [x for x in rows if key(x["firma"]) and
                   (key(x["firma"]) in fk or fk in key(x["firma"]))]
        if len(treffer) == 1:
            return treffer[0]
        if len(treffer) > 1 and pk:
            genau = [x for x in treffer if poskey(x["pos"]) and
                     (poskey(x["pos"])[:20] in pk or pk[:20] in poskey(x["pos"]))]
            if genau:
                return genau[0]
            return treffer[0]
        if treffer:
            return treffer[0]
    if pk:
        treffer = [x for x in rows if poskey(x["pos"]) and
                   (poskey(x["pos"])[:24] in pk or pk[:24] in poskey(x["pos"]))]
        if len(treffer) == 1:
            return treffer[0]
    return None


RANG = {"Entwurf unvollständig": 0, "Vorbereitet": 0, "Beworben": 1,
        "Eingangsbestätigung": 2, "Coding-Challenge": 3, "Gespräch": 3,
        "Gespräch geführt": 4, "Absage": 5, "Zusage": 5, "Verworfen": 5}


def next_row(ws):
    r = ws.max_row + 1
    # in die erste wirklich leere Zeile, damit Leerzeilen genutzt werden
    for cand in range(FIRST_DATA_ROW, ws.max_row + 2):
        if not any(ws.cell(row=cand, column=c).value
                   for c in (C_DATUM, C_FIRMA, C_POS, C_STATUS)):
            r = cand
            break
    return r


def write_new(ws, vorgang):
    r = next_row(ws)
    nr = max([ws.cell(row=x, column=C_NR).value or 0
              for x in range(FIRST_DATA_ROW, r)] or [0]) + 1
    thin = Side(style="thin", color="E7ECEF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    werte = {
        C_NR: nr, C_DATUM: vorgang["datum"], C_FIRMA: vorgang["firma"],
        C_POS: vorgang["pos"], C_QUELLE: vorgang["quelle"],
        C_STATUS: vorgang["status"], C_AKTION: vorgang["datum"],
        C_NOTIZ: vorgang["notiz"],
    }
    for c in range(1, N_COLS + 1):
        cell = ws.cell(row=r, column=c, value=werte.get(c))
        cell.font = Font(name="Arial", size=10)
        cell.alignment = Alignment(vertical="top",
                                   wrap_text=(c in (C_POS, C_SCHRITT, C_NOTIZ)))
        cell.border = border
    for c in (C_DATUM, C_AKTION, C_WV):
        ws.cell(row=r, column=c).number_format = "DD.MM.YYYY"
    ws.row_dimensions[r].height = 32
    return r


# =========================================================================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true",
                    help="nur zeigen, nichts schreiben")
    ap.add_argument("--since", help="ab diesem Datum suchen (JJJJ-MM-TT)")
    ap.add_argument("--xlsx", default=str(XLSX))
    args = ap.parse_args()

    pfad = Path(args.xlsx)
    neu_angelegt = False
    if not pfad.exists():
        # Fehlt die Datei, wird sie angelegt statt abzubrechen. Sonst
        # steht die ganze Kette, nur weil jemand die Tabelle gelöscht hat.
        if args.dry_run:
            print(f"{pfad} fehlt. Ein echter Lauf würde sie anlegen.")
        else:
            wb_neu = Workbook()
            wb_neu.remove(wb_neu.active)
            wb_neu.create_sheet("Legende")
            ws_lg = wb_neu["Legende"]
            ws_lg.column_dimensions["A"].width = 86
            ws_lg.cell(row=1, column=1,
                       value="Von mailcheck.py neu angelegt am "
                             f"{dt.date.today().strftime('%d.%m.%Y')}. "
                             "Ein Blatt pro Monat, Status in Spalte I.")
            wb_neu.save(pfad)
            neu_angelegt = True
            print(f"{pfad.name} war nicht vorhanden und wurde neu angelegt.")

    state = load_state()
    if args.since:
        seit = dt.date.fromisoformat(args.since)
    elif state.get("last_run"):
        seit = dt.date.fromisoformat(state["last_run"]) - dt.timedelta(days=2)
    else:
        seit = dt.date.today() - dt.timedelta(days=14)
    gesehen = set(state.get("seen", []))

    user, pw = creds()
    print(f"Postfach {user}, Mails ab {seit.strftime('%d.%m.%Y')}")
    try:
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        imap.login(user, pw)
    except imaplib.IMAP4.error as e:
        sys.exit(f"Anmeldung fehlgeschlagen: {e}\n"
                 f"App-Passwort prüfen (nicht das normale Kontopasswort).")

    nachrichten = []
    for folder, quelle in (("INBOX", "eingehend"),
                           (sent_folder(imap), "gesendet")):
        if not folder:
            continue
        msgs = fetch_folder(imap, folder, seit)
        print(f"  {folder}: {len(msgs)} Mails")
        nachrichten += [(m, quelle) for m in msgs]
    imap.logout()

    vorgaenge, neue_ids = [], []
    for msg, richtung in nachrichten:
        mid = decode(msg.get("Message-ID")) or ""
        if mid and mid in gesehen:
            continue
        subject = decode(msg.get("Subject"))
        sender = decode(msg.get("From"))
        if richtung == "eingehend" and noise(sender, subject):
            neue_ids.append(mid)
            continue
        body = body_text(msg)

        if richtung == "gesendet":
            # nur eigene Bewerbungsmails, nicht jede Antwort
            if not re.search(r"bewerbung|application", subject, re.I):
                neue_ids.append(mid)
                continue
            if re.match(r"\s*(re|aw|antwort|fwd|wg)\s*:", subject, re.I):
                neue_ids.append(mid)
                continue
            status, grund = "Beworben", "selbst verschickt"
            empf = decode(msg.get("To"))
            firma = firma_aus_text(subject, body, empf)
        else:
            status, grund = classify(subject, body)
            if not status:
                neue_ids.append(mid)
                continue
            firma = firma_aus_text(subject, body, sender)

        try:
            datum = parsedate_to_datetime(msg.get("Date")).date()
        except (TypeError, ValueError):
            datum = dt.date.today()

        vorgaenge.append({
            "datum": datum, "firma": firma,
            "pos": position_aus_text(subject, body),
            "status": status, "subject": subject, "sender": sender,
            "quelle": "Mail" if richtung == "gesendet" else "Portal",
            "grund": grund, "mid": mid, "richtung": richtung,
        })
        neue_ids.append(mid)

    vorgaenge.sort(key=lambda v: v["datum"])
    print(f"\n{len(vorgaenge)} relevante Mails\n")

    wb = load_workbook(pfad)
    rows = read_rows(wb)
    geaendert_liste, beruehrt, angelegt, unklar = [], [], [], []

    for v in vorgaenge:
        treffer = find_row(rows, v["firma"], v["pos"])
        kennung = f"{v['datum'].strftime('%d.%m.')} {v['firma'] or '?'} · {v['status']}"

        if treffer:
            ws = wb[treffer["sheet"]]
            alt = treffer["status"]
            # Nur bei echtem Fortschritt überschreiben. Bei gleichem Rang
            # bleibt der genauere Eintrag stehen: 'Coding-Challenge' soll
            # nicht durch das allgemeinere 'Gespräch' ersetzt werden.
            geaendert = RANG.get(v["status"], 1) > RANG.get(alt, 0)
            if geaendert:
                if not args.dry_run:
                    ws.cell(row=treffer["row"], column=C_STATUS, value=v["status"])
                treffer["status"] = v["status"]
            if not args.dry_run:
                ws.cell(row=treffer["row"], column=C_AKTION, value=v["datum"])
                ws.cell(row=treffer["row"], column=C_AKTION).number_format = "DD.MM.YYYY"
                alt_notiz = ws.cell(row=treffer["row"], column=C_NOTIZ).value or ""
                zusatz = f"{v['datum'].strftime('%d.%m.')}: {v['subject'][:90]}"
                if zusatz[:14] not in alt_notiz:
                    ws.cell(row=treffer["row"], column=C_NOTIZ,
                            value=(alt_notiz + " | " + zusatz).strip(" |")[:900])
            if geaendert:
                geaendert_liste.append(f"{kennung}  (war: {alt})")
            else:
                beruehrt.append(f"{v['datum'].strftime('%d.%m.')} "
                                f"{v['firma'] or '?'} · bleibt {alt} "
                                f"(Mail deutete auf {v['status']})")
            continue

        if v["status"] in ("Eingangsbestätigung", "Beworben") and v["firma"]:
            ym = v["datum"].strftime("%Y-%m")
            ws, neu = ensure_month(wb, ym)
            if neu:
                print(f"  → neues Blatt {ym} angelegt")
            v["notiz"] = f"Automatisch erkannt: {v['subject'][:120]}"
            if not args.dry_run:
                r = write_new(ws, v)
                rows.append({"sheet": ym, "row": r, "firma": v["firma"],
                             "pos": v["pos"], "status": v["status"]})
            angelegt.append(f"{kennung} · {v['pos'] or 'Position unklar'}")
        else:
            unklar.append(f"{kennung} · {v['subject'][:70]}")

    if not args.dry_run:
        wb.save(pfad)

    def block(titel, items):
        if items:
            print(f"{titel} ({len(items)}):")
            for x in items:
                print(f"  · {x}")
            print()

    block("Status fortgeschrieben", geaendert_liste)
    block("Datum aktualisiert, Status unverändert", beruehrt)
    block("Neu angelegt", angelegt)
    block("Nicht zugeordnet, bitte von Hand prüfen", unklar)

    if args.dry_run:
        print("Probelauf — die Tabelle wurde nicht verändert.")
    else:
        STATE.parent.mkdir(parents=True, exist_ok=True)
        STATE.write_text(json.dumps({
            "last_run": dt.date.today().isoformat(),
            "seen": (list(gesehen) + neue_ids)[-4000:],
        }))
        print(f"{pfad.name} gespeichert."
              + (" Die Blätter Übersicht und Legende musst du bei einer"
                 " neu angelegten Datei selbst ergänzen." if neu_angelegt else ""))


if __name__ == "__main__":
    sys.exit(main())
# job-search-automation

Zwei kleine Python-Dienste, die meine eigene Stellensuche automatisieren. Sie
laufen täglich, suchen, filtern, bewerten und schreiben die Bewerbungstabelle
fort. Beworben wird weiterhin von Hand.

## Warum

Eine Stellensuche ist zu großen Teilen wiederkehrende Datenarbeit: dieselben
Portale abfragen, dieselben Filter anlegen, und bei jedem Treffer die Frage,
ob man sich dort schon beworben hat. Statt das jeden Tag von Hand zu machen,
habe ich es einmal aufgeschrieben.

Das Projekt ist aus dem Bedarf entstanden, nicht als Übung. Es läuft seit
Wochen im Alltag und wird geändert, wenn im Betrieb etwas nicht stimmt.

## Die beiden Dienste

### `jobfinder.py`

Fragt die offizielle [Jobsuche-API der Bundesagentur für
Arbeit](https://github.com/bundesAPI/jobsuche-api) mit rund zwanzig
Suchprofilen ab, etwa 700 Treffer pro Lauf.

* filtert Fehltreffer über Titel- und Angebotsart-Sperrlisten
* prüft das Veröffentlichungsdatum selbst, weil der Serverfilter der API
  nicht zuverlässig greift
* prüft die Entfernung selbst, aus demselben Grund
* gleicht gegen die Bewerbungstabelle ab und erkennt, ob dort schon eine
  Bewerbung beim selben Arbeitgeber läuft
* bewertet jeden Treffer gewichtet und sortiert die Tabelle danach
* schreibt in eine formatierte `.xlsx` mit Statusauswahl, bedingter
  Formatierung und Autofilter

### `mailcheck.py`

Liest das Postfach per IMAP und schreibt die Bewerbungstabelle fort.

* erkennt Eingangsbestätigungen, Absagen, Einladungen und Erinnerungen an
  unvollständige Bewerbungen
* durchsucht auch den Gesendet-Ordner nach Bewerbungen, die direkt per Mail
  verschickt wurden
* liest Arbeitgeber und Position aus dem Fließtext
* findet die passende Zeile und aktualisiert Status, letzte Aktion und Notiz
* legt für einen neuen Monat selbst ein neues Blatt an

## Zwei Entscheidungen, die mehr Arbeit gemacht haben als der Rest

### Dubletten über Inhaltswörter statt Zeichenvergleich

Ein Ähnlichkeitsmaß auf Zeichenebene trennt Stellentitel nicht sauber.
*Innovation Garage IoT Backend Development* und *Innovation Garage Cyber
Security* liegen bei 0,82 — derselbe IoT-Titel anders formuliert bei 0,88.
Drei Hundertstel Abstand zwischen „zwei verschiedene Stellen" und „dieselbe
Stelle, anders geschrieben" ist keine brauchbare Grenze.

Deshalb wird über die Inhaltswörter verglichen, geteilt durch die kleinere
Wortmenge, damit ein verkürzter Titel noch zuordenbar bleibt. Füllwörter wie
*Werkstudent*, *m/w/d* oder *Schwerpunkt* werden vorher entfernt, weil sie in
fast jedem Titel stehen und nichts unterscheiden.

### Bewertung ohne Modell

Die Passung wird über ein gewichtetes Punkteschema auf den Stellentitel
bestimmt. Konkrete Technologien wiegen schwer, generische Begriffe wie
*IT* oder *System* kaum. Dazu kommen Zuschläge für Nähe, Homeoffice,
Vertragsdauer und Frische.

Entscheidend war eine Begrenzung: **Die Zuschläge sind auf die Höhe des
fachlichen Kerns gedeckelt.** Ohne das stieg eine IT-Support-Stelle allein
dadurch auf „hoch", dass sie zehn Kilometer entfernt und fünf Tage alt war.
Ein Titel ganz ohne Fachbezug bekommt null Punkte, egal wie nah oder frisch.

Regelbasiert statt gelernt, und das bewusst: Bei einigen hundert Einträgen ist
ein nachvollziehbares Punkteschema belastbarer als etwas Trainiertes, und es
lässt sich jederzeit erklären, warum ein Eintrag oben steht.

## Was die Dienste bewusst nicht tun

Sie bewerben sich nicht, beantworten keine Mail und verschicken nichts.
`mailcheck.py` liest nur und schreibt in die Tabelle, `jobfinder.py` sucht und
trägt ein. Jede Bewerbung geht weiterhin von Hand raus.

## Technik

Python 3, `requests`, `openpyxl`. Kein Framework und keine Datenbank: Das
Ergebnis ist eine Tabellendatei, weil ich damit ohnehin arbeite.

Nicht im Repository sind die Tabellen selbst, der Zustandsspeicher und das
Log — die enthalten laufende Bewerbungen und Mail-Historie.

## Einrichtung

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

`mailcheck.py` meldet sich per IMAP bei Gmail an und braucht ein
**App-Passwort**, nicht das Kontopasswort:

```bash
cat > ~/.jobsuche-mail.json <<'EOF'
{"user": "deine.adresse@gmail.com", "password": "APP-PASSWORT"}
EOF
chmod 600 ~/.jobsuche-mail.json
```

Alternativ über die Umgebungsvariablen `GMAIL_USER` und
`GMAIL_APP_PASSWORD`.

## Benutzung

```bash
./tagesstart.sh | tee -a log.txt          # beides nacheinander

.venv/bin/python mailcheck.py --dry-run   # zeigt nur, was es ändern würde
.venv/bin/python jobfinder.py             # nach Passung sortiert
.venv/bin/python jobfinder.py --sortierung datum
.venv/bin/python jobfinder.py --debug     # Rohfelder des ersten Treffers
```

Die Suchprofile stehen oben in `jobfinder.py` in der Liste `SEARCHES`, die
Gewichte darunter in `TECH_PUNKTE`. Nach einer Änderung wird beim nächsten
Lauf die gesamte Tabelle neu bewertet, die Wirkung ist also sofort sichtbar.

## Lizenz

MIT

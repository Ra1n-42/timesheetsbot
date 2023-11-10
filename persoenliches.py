import os
import glob
import subprocess


from datetime import datetime, date, timedelta


def save_pdf(excel, pdf):

    try:
        fertige_pdf = glob.glob(OUTPUT_PATH + "*")[0]

        if os.path.exists(fertige_pdf):
            os.remove(fertige_pdf)
    except:pass

    libreoffice_path = 'C:/Program Files/LibreOffice/program/scalc.exe'  # Beispiel für Windows
    subprocess.call([libreoffice_path, '--headless', '--convert-to', 'pdf', '--outdir', os.path.dirname(pdf), excel])
    name = pdf.split("/")
    output_filename = name[len(name)-1].split(".")[0]

    # Benennen der PDF-Datei

    os.rename(os.path.join(OUTPUT_PATH, os.path.basename(excel).split('.')[0] + ".pdf"),
              os.path.join(OUTPUT_PATH, output_filename + ".pdf"))


SEND_MAIL_FROM = "muster@mann.de" # Sender
SEND_MAIL_TO = "muster@frau.de"  # Empfänger

VORNAME = "Vorname"
NACHNAME = "Nachname"
STRASSE = "Straße 007"
ORT = "Der Ort"
AUFTRAGGEBER = "Selbserklärend"
ORT_BAUSTELLE = "Selbserklärend"
PERSONAL_NR = ["", "", "", "", ""]
BAUSTELLEN = {
    # Auftraggeber : [Kostenstelle]
    'Auftraggeber': [1, 1, 3, 8, 0, 8, 0, 0],
    'DB-AKADEMIE': ["", "", "", "", "", "", "", ""],
}

# Pfad der Vorlagedatei
WORKBOOK_PATH = 'D:/PycharmProject/vorlage2/'+AUFTRAGGEBER+'-VORLAGE.xlsx'

OUTPUT_PATH = 'D:/PycharmProject/readytosend/'

TODAY = date.today()
YEAR, WEEKNUMBER, day_of_week = TODAY.isocalendar()

# Berechnung des ersten Tages der Kalenderwoche
FIRST_DAY = datetime.strptime(f'{YEAR}-W{WEEKNUMBER}-1', '%G-W%V-%u').date()

# Berechnung des letzten Tages der Kalenderwoche
LAST_DAY = FIRST_DAY + timedelta(days=6)

# KW als 2 digit-string
KALENDERWOCHE = str(WEEKNUMBER).zfill(2)

DATEN = {
    'K9': FIRST_DAY,
    'AN9': LAST_DAY,
    'K11': NACHNAME,
    'AP11': VORNAME,
    'BJ11': AUFTRAGGEBER,
    'BJ13': ORT_BAUSTELLE,
    # Personal-Nr.
    'N21': PERSONAL_NR[0],
    'Q21': PERSONAL_NR[1],
    'T21': PERSONAL_NR[2],
    'W21': PERSONAL_NR[2],
    'Z21': PERSONAL_NR[2],
    # KW
    'AD21': KALENDERWOCHE[0],
    'AG21': KALENDERWOCHE[1],
    # Baustelle Nummer
    'D25': BAUSTELLEN[AUFTRAGGEBER][0],
    'F25': BAUSTELLEN[AUFTRAGGEBER][1],
    'H25': BAUSTELLEN[AUFTRAGGEBER][2],
    'J25': BAUSTELLEN[AUFTRAGGEBER][3],
    'L25': BAUSTELLEN[AUFTRAGGEBER][4],
    'N25': BAUSTELLEN[AUFTRAGGEBER][5],
    'P25': BAUSTELLEN[AUFTRAGGEBER][6],
    'R25': BAUSTELLEN[AUFTRAGGEBER][7],

    # AB
    'Z25': 8, 'AD25': 75,   # Montag
    'AF25': 8, 'AJ25': 75,  # Dienstag
    'AL25': 8, 'AP25': 75,  # Mittwoch
    'AR25': 8, 'AV25': 75,  # Donnerstag
    'AX25': 6, 'BB25': 0,   # Freitag

    # Arbeitszeit     von         -          bis
    'Z36':     '06:30', 'Z38':   '16:00',   # Montag
    'AF36':    '06:30', 'AF38':  '16:00',   # Dienstag
    'AL36':    '06:30', 'AL38':  '16:00',   # Mittwoch
    'AR36':    '06:30', 'AR38':  '16:00',   # Donnerstag
    'AX36':    '06:30', 'AX38':  '13:00',   # Freitag

    'BY88': 'X',
    'Z103': 'X',
    'BN103': 'X',

}

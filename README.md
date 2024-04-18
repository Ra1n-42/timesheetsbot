# Wofür der Bot? (nur für Windows + LibreOffice)
Warum nicht ein Skript schreiben, das den Stundenzettel in weniger als einer Sekunde erstellt und ihn per E-Mail verschickt?


# Was man einstellen muss:

Die Konfigurationsdatei wird als Modul geladen und heißt "persoenliches.py", dort werden sämtliche Variablen festgelegt.

1. Hier wird der Pfad zur Scalc aka LibreOffice festgelegt.
   
````
libreoffice_path = 'C:/Program Files/LibreOffice/program/scalc.exe'  # Beispiel für Windows
````

2. Im "WORKBOOK_PATH" wird ein Vorlagenpfad benötigt, der sich aus dem Namen des Auftraggebers und "-VORLAGE.xlsx" zusammensetzt.
   Dies wird verwendet, um später eine Datei zu erstellen, die gesendet wird.
   Der "OUTPUT_PATH" ist der Pfad zu einem Ordner, wo die Datei zum Senden erstellt wird (sagt auch der Name :D).
````
WORKBOOK_PATH = 'D:/PycharmProject/vorlage/'+AUFTRAGGEBER+'-VORLAGE.xlsx'
OUTPUT_PATH = 'D:/PycharmProject/readytosend/'
````

3. In der "DATEN" Dictionary werden die Adressen der Zellen hinterlegt. Diese können ganz einfach editiert werden, falls sich das Format in der Vorlagendatei ändert.

# Zusatz!
Im jetzigen Stand wird nur die fertige Datei zum Senden erstellt. Um die Datei via Mail zu senden, entfernen Sie einfach die # Zeichen in der "main.py".
````
    # mailing.send_mail(SEND_MAIL_FROM, SEND_MAIL_TO, f'Stundenzettel KW{WEEKNUMBER} {NACHNAME}, {YEAR}',
    #                   f"""Hier ist der Stundenzettel für KW{WEEKNUMBER}.\n\nLiebe grüße, Rain's Bot."""
    #                   , files=output_path)
````


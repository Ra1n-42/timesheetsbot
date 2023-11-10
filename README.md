# Wofür der Bot? (nur für Windows + LibreOffice)
Waum nicht ein skript schreiben, der den Stundenzettel in weniger als einer Sekunde erstellt und es via Mail verschickt.


# Was man einstellen muss.

Die Configfile wird als modul geladen und heißt persoenliches.py, dort werden sämmtliche variablen festgelegt.

1. Hier wird der Pfad zur scalc aka LibreOffice festgelegt.
   
````
libreoffice_path = 'C:/Program Files/LibreOffice/program/scalc.exe'  # Beispiel für Windows
````

2. in den WORKBOOK_PATH Pfad wird eine vorlage benötigt, setzt sich zusammen aus Name des AUFTRAGGEBER's und '-VORLAGE.xlsx'
   diese wird verwendet um später eine Datei zu erstellen, die gesendet wird.
   OUTPUT_PATH ist der Pfad zu einen ordner, wo die Datei zum senden erstellt wird (sagt auch der Name :D)
````
WORKBOOK_PATH = 'D:/PycharmProject/vorlage/'+AUFTRAGGEBER+'-VORLAGE.xlsx'
OUTPUT_PATH = 'D:/PycharmProject/readytosend/'
````

3. In der DATEN Dictionary werden die Adressen der Zellen hinterlegt, diese können ganz einfach editiert werden, falls sich das Format in der Volagendatei ändert.

# Was brauche ich um den Bot zu starten?
1. Python 3.9 als Standalone
2. Die main.py auf die Exe ziehen und viel Spaß.

# Zusatz!
Im jetzigen Stand, wird nur die Fertige Datei zum senden erstellt, um die Datei via Mail zu senden einfach in der main.py 
````
    # mailing.send_mail(SEND_MAIL_FROM, SEND_MAIL_TO, f'Stundenzettel KW{WEEKNUMBER} {NACHNAME}, {YEAR}',
    #                   f"""Hier ist der Stundenzettel für KW{WEEKNUMBER}.\n\nLiebe grüße, Rain's Bot."""
    #                   , files=output_path)
````

die # Zeichen entfernen.

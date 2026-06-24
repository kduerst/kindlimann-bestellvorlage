# Kindlimann Bestellliste - Entwicklungsstruktur

## Fertige Datei

Die testbare Excel-Datei liegt im Ordner darüber:

`C:\Meine Programme\Kindlimann\Kindlimann Bestellliste.xlsm`

## Konfiguration

Pflege einfache Werte hier:

`Config\kindlimann.settings.psd1`

Dort liegen:

- Zielordner und Dateiname
- Logo-Pfad
- Schutzpasswort
- sichtbarer Start ab KW27
- Profil-Dropdownliste
- E-Mail-Empfänger und Mailtext-Bausteine
- Layoutwerte wie Spaltenbreiten, Zeilenhöhen, Button-Abstand und PDF-Rand

## Generator

`create_kindlimann_workbook.ps1` erzeugt die Excel-Datei aus der Konfiguration und bettet die Makros ein.

`run_generator_with_vbom.ps1` startet den Generator und setzt dafür temporär die Excel-Einstellung, damit VBA-Code eingefügt werden darf.

## Wichtig

Kleine Änderungen an Profilen, Mailtext, Rand oder Abständen zuerst in der Config machen. Den Generator nur ändern, wenn sich die eigentliche Logik ändern soll.

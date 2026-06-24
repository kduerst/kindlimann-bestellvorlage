# Ist-Zustand vor Ausfuehrung der Arbeitsauftraege

## Zeitpunkt

24. Juni 2026, 19:10 Uhr

## Repository

`kduerst/kindlimann-bestellvorlage`

## Zweck dieses Dokuments

Dieses Dokument haelt den bekannten Stand fest, bevor die erfassten Arbeitsauftraege ausgefuehrt werden.

Es dient als Kontrollpunkt, damit spaeter nachvollziehbar bleibt:

- was aktuell vorhanden ist
- was laut bisherigem Wissensstand funktionieren soll
- welche Auftraege bereits formuliert wurden
- welche Bereiche ausdruecklich noch nicht veraendert werden duerfen

Dieses Dokument ist eine Bestandsaufnahme. Es fuehrt keine Codeaenderung aus.

## Geltungsbereich

Dieses Dokument gilt ausschliesslich fuer das Repository:

```text
kduerst/kindlimann-bestellvorlage
```

Andere Repositories dienen nicht als fachliche Grundlage. Die Arbeitsweise mit Arbeitsauftraegen und Entscheidungen wurde nur als Strukturprinzip uebernommen.

## Bekannter aktueller Stand

Die Kindlimann Bestellliste ist als Excel-Arbeitsmappe vorhanden und funktioniert nach bisheriger Rueckmeldung bereits grundsaetzlich gut.

Die Datei dient zur woechentlichen Sammelbestellung von Kontraktmaterial bei Kindlimann.

Bestelllogik laut bisherigem Projektstand:

- Bestellungen werden jeweils dienstags vorbereitet.
- Gewuenschter Liefertermin ist jeweils der folgende Donnerstag.
- Pro Kalenderwoche koennen mehrere Auftraege oder Projekte erfasst werden.
- Die Bestellung soll am Ende als PDF abgelegt und per Mail versendet werden koennen.

## Aktuell bekannte Repository-Bestandteile

Aus der aktuellen Dokumentation und Konfiguration sind folgende Bestandteile bekannt:

```text
Kindlimann Bestellliste.xlsm
Bestellt/
_Entwicklung/
_Entwicklung/Config/kindlimann.settings.psd1
README.md
.gitignore
```

Die produktive Datei ist laut README:

```text
Kindlimann Bestellliste.xlsm
```

Der Ordner `Bestellt` ist als Zielordner fuer erzeugte PDFs vorgesehen. Exportierte PDFs werden nicht versioniert.

## Bekannte zentrale Konfiguration

Die Datei `_Entwicklung/Config/kindlimann.settings.psd1` enthaelt nach aktuellem Stand zentrale Einstellungen, unter anderem:

- Zielordner
- Ausgabedateiname
- Logo-Pfad
- Passwort
- Mindest-Startdatum
- Profil-Liste
- Mail-Einstellungen
- Layout-Einstellungen

## Hinterlegte Profile

Aktuell sollen folgende Profile fuer die Kindlimann-Bestellung hinterlegt sein:

- 40 x 15 x 2 mm
- FL 30 x 6 mm
- FL 40 x 10 mm
- FL 40 x 15 mm
- RD 10 mm

## Bekannte Maildaten

Die Konfiguration enthaelt Mailangaben fuer den Versand der Sammelbestellung.

Bekannte Werte:

- Empfaenger: claudia.desantis@kindlimann.ch
- Anrede: Guten Tag Frau De Santis
- Betreff-Prefix: Sammelbestellung Metallraum
- Grussformel: Freundliche Gruesse
- Signatur: Metallraum

Diese Werte duerfen ohne eigenen Arbeitsauftrag nicht geaendert werden.

## Bekannte Layout-Parameter

In der Konfiguration sind Layoutwerte hinterlegt, unter anderem:

- Button-Abstand
- PDF-Rand
- Spaltenbreiten
- Zeilenhoehen

Diese Werte sind fuer spaetere Layout-Arbeiten relevant, duerfen aber im Rahmen dieser Bestandsaufnahme nicht geaendert werden.

## Aktuell bekannte Funktionen, die funktionieren sollen

Folgende Funktionen sind nach bisherigem Projektverlauf gewuenscht oder bereits vorhanden und sollen bei spaeteren Arbeiten geprueft werden:

- Sammelbestellung pro KW
- mehrere Auftraege pro KW
- Besteller-Auswahl per Dropdown
- Profil-Auswahl per Dropdown
- Pluszeichen bei Auftraegen zum Hinzufuegen von Profilen
- automatisches Bestelldatum / Druckdatum
- automatischer Liefertermin
- PDF-Erstellung / PDF-Export
- Mailvorbereitung oder Mailversand
- Ablage erzeugter PDFs im Ordner `Bestellt`
- Ausschluss erzeugter PDFs aus Git

## Bekannte Git-Regeln

Die `.gitignore` regelt nach aktuellem Stand insbesondere:

- Office-Temporaerdateien werden ignoriert
- exportierte PDFs unter `Bestellt/*.pdf` werden nicht versioniert
- Logdateien werden ignoriert

Der Ordner `Bestellt` soll erhalten bleiben, ohne die erzeugten PDFs hochzuladen.

## Bereits formulierter Arbeitsauftrag 001

Arbeitsauftrag 001 lautet sinngemaess:

Die Repository-Struktur der Kindlimann Bestellliste soll geordnet werden. Ziel ist eine saubere, nachvollziehbare und wartbare Ablage, damit zukuenftige kleine Aenderungen gezielt und mit geringem Risiko erfolgen koennen.

Wichtige Vorgaben:

- keine Aenderung am VBA-/Makro-Code
- keine Aenderung an der produktiven Excel-Logik
- keine Aenderung an bestehenden Formeln
- keine Aenderung am PDF-Export-Verhalten
- keine Aenderung an Mailtexten oder Empfaengern
- keine Aenderung an hinterlegten Profilen
- keine optische Ueberarbeitung der Excel-Datei
- keine funktionale Erweiterung

Erwartetes Ergebnis:

```text
Dokumentation/Entscheidungen/Arbeitsauftrag_001_Repository_Struktur_Kindlimann_Bestellliste.md
```

## Bereits formulierter Arbeitsauftrag 002

Arbeitsauftrag 002 darf erst nach erfolgreichem Abschluss und Freigabe von Arbeitsauftrag 001 ausgefuehrt werden.

Thema:

Benutzeroberflaeche, Layout und Bedienpruefung.

Wichtige Punkte:

- ca. 20 mm Abstand rund um die Tabelle
- Tabelle soll nicht oben links in der Ecke kleben
- Spaltenbreiten und Zeilenhoehen passend zum Inhalt ausrichten
- nicht gequetscht, nicht unnoetig gross
- Pluszeichen-Funktion pruefen
- Profil-Dropdowns pruefen
- Besteller-Dropdown pruefen
- Pflichtfelder und Kontrollanzeigen pruefen
- PDF-Export nach Layoutanpassung pruefen
- Farbgestaltung schlicht, stimmig und professionell
- Code nur so lang wie noetig halten
- keine benoetigten Funktionen entfernen

Erwartetes Ergebnis:

```text
Dokumentation/Entscheidungen/Arbeitsauftrag_002_Benutzeroberflaeche_Layout_Pruefung.md
```

## Offene Unsicherheiten

Folgende Punkte sind aus diesem Chat heraus noch nicht technisch verifiziert:

- exakter aktueller VBA-Code
- exakte aktuelle Blattstruktur in der `.xlsm`
- ob der PDF-Export aktuell fehlerfrei in Excel Desktop laeuft
- ob die Pluszeichen-Funktion in jeder relevanten Situation korrekt arbeitet
- ob alle Dropdowns in der produktiven Datei korrekt gesetzt sind
- ob die Layoutwerte bereits in allen KW-/Bestellbereichen konsistent wirken

Diese Punkte sollen nicht geraten, sondern in den passenden Arbeitsauftraegen geprueft und dokumentiert werden.

## Entscheidung fuer das weitere Vorgehen

Vor jeder technischen Umsetzung muss zuerst der Ist-Zustand verstanden und dokumentiert werden.

Die bisherigen Auftraege definieren den Wunschzustand. Dieses Dokument definiert den bekannten Ausgangspunkt.

Arbeitsauftrag 001 ist der erste auszufuehrende Auftrag.
Arbeitsauftrag 002 darf erst nach Arbeitsauftrag 001 ausgefuehrt werden.

## Keine Codeaenderung

Mit diesem Dokument wurden keine Code-, Excel-, VBA-, Layout- oder Funktionsaenderungen verlangt oder vorgenommen.

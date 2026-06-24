# Arbeitsauftrag 001 - Ist-Zustand vor Umsetzung aufnehmen

## Repository

`kduerst/kindlimann-bestellvorlage`

## Zeitpunkt

24. Juni 2026, 19:10 Uhr

## Ziel

Vor jeder technischen Umsetzung soll zuerst der aktuelle Ist-Zustand der Kindlimann Bestellliste verstanden, dokumentiert und bewertet werden.

Der Wunschzustand wird durch spaetere Arbeitsauftraege definiert. Dieser erste Auftrag dient dazu, den Ausgangspunkt festzuhalten, damit spaeter kontrolliert werden kann, ob die Umsetzung tatsaechlich den gewuenschten Anforderungen entspricht.

## Grundsatz

Dieser Arbeitsauftrag ist ein reiner Analyse- und Dokumentationsauftrag.

Es duerfen keine Code-, Excel-, VBA-, Layout- oder Funktionsaenderungen vorgenommen werden.

## Geltungsbereich

Dieser Auftrag gilt ausschliesslich fuer:

```text
kduerst/kindlimann-bestellvorlage
```

Andere Repositories dienen nicht als fachliche Grundlage. Die Arbeitsweise mit Arbeitsauftraegen und Entscheidungen kann als Strukturprinzip genutzt werden.

## Bekannter aktueller Stand

Die Kindlimann Bestellliste ist als Excel-Arbeitsmappe vorhanden und funktioniert nach bisheriger Rueckmeldung grundsaetzlich gut.

Die Datei dient zur woechentlichen Sammelbestellung von Kontraktmaterial bei Kindlimann.

Bekannte Bestelllogik:

- Bestellungen werden jeweils dienstags vorbereitet.
- Gewuenschter Liefertermin ist der folgende Donnerstag.
- Pro Kalenderwoche koennen mehrere Auftraege oder Projekte erfasst werden.
- Die Bestellung soll am Ende als PDF abgelegt und per Mail versendet werden koennen.

## Bekannte Repository-Bestandteile

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

Der Ordner `Bestellt` ist als Zielordner fuer erzeugte PDFs vorgesehen. Exportierte PDFs sollen nicht versioniert werden.

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

Diese Werte sind fuer spaetere Layout-Arbeiten relevant, duerfen aber in diesem Auftrag nicht geaendert werden.

## Aktuell bekannte Funktionen, die funktionieren sollen

Folgende Funktionen sind nach bisherigem Projektverlauf gewuenscht oder bereits vorhanden und sollen im Ist-Zustand geprueft werden:

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

Der Ordner `Bestellt` soll erhalten bleiben, ohne erzeugte PDFs hochzuladen.

## Zu pruefende offene Punkte

Folgende Punkte sind noch nicht technisch verifiziert und sollen im Rahmen dieses Arbeitsauftrags geprueft oder als offen dokumentiert werden:

- exakter aktueller VBA-Code
- exakte aktuelle Blattstruktur in der `.xlsm`
- ob der PDF-Export aktuell fehlerfrei in Excel Desktop laeuft
- ob die Pluszeichen-Funktion in jeder relevanten Situation korrekt arbeitet
- ob alle Dropdowns in der produktiven Datei korrekt gesetzt sind
- ob die Layoutwerte bereits in allen KW-/Bestellbereichen konsistent wirken

## Vorgehen

1. Aktuelle Repository-Struktur aufnehmen.
2. Vorhandene Dokumente und Konfigurationen lesen.
3. Bekannte Funktionen erfassen.
4. Unsichere oder nicht verifizierte Punkte sauber markieren.
5. Keine Aenderung an Code, Excel, VBA oder Layout vornehmen.
6. Ergebnis als Entscheidungsdokument ablegen.

## Erwartetes Ergebnisdokument

Nach Abschluss ist folgendes Dokument zu erstellen:

```text
Dokumentation/Entscheidungen/Arbeitsauftrag_001_Ist_Zustand_vor_Umsetzung.md
```

Dieses Ergebnisdokument muss enthalten:

- Ziel
- Vorgehen
- festgestellter Ist-Zustand
- bekannte Funktionen
- bekannte Konfigurationen
- nicht verifizierte Punkte
- Risiken
- Entscheidung
- Empfehlung fuer den naechsten Arbeitsauftrag

## Freigabe

Die finale Freigabe erfolgt durch Kevin.

Arbeitsauftrag 002 darf erst begonnen werden, wenn dieser Ist-Zustand dokumentiert und freigegeben ist.

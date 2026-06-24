# Arbeitsauftrag 000 - Machbarkeit HiCAD-Profilstabuebernahme in Bestellliste

## Repository

`kduerst/kindlimann-bestellvorlage`

## Einordnung

Dieser Arbeitsauftrag wird bewusst vor Arbeitsauftrag 001 eingeordnet.

Grund: Bevor die bestehende Kindlimann Bestellliste strukturell weiterentwickelt wird, soll fachlich und technisch verstanden werden, ob und wie eine spaetere Verbindung zur HiCAD-Stueckliste beziehungsweise Profilstabzusammenfassung sinnvoll moeglich ist.

## Ziel

Pruefen und dokumentieren, wie eine spaetere Funktion aussehen koennte, mit der Profile aus einer HiCAD-nahen Excel-Tabelle oder Profilstabzusammenfassung per Kontrollkaestchen ausgewaehlt und in die Kindlimann Bestellliste uebertragen werden.

Der Fokus liegt auf Machbarkeit, Datenfluss, Risiken und sinnvoller Zielarchitektur.

## Ausgangsidee

In HiCAD beziehungsweise in der zugehoerigen Excel-/Stuecklisten-Ausgabe gibt es eine Profilstabzusammenfassung.

Diese Profilstabzusammenfassung zeigt Profile projektbezogen auf 6-Meter-Stangen gerechnet an. Beispiele:

- 4 Stangen
- 5 Stangen
- 100 Stangen

Gewuenscht ist eine spaetere Bedienlogik:

1. In der Profilstabzusammenfassung gibt es pro Profilzeile ein Kontrollkaestchen.
2. Der Benutzer waehlt die gewuenschten Profile aus.
3. Ein Button startet die Uebernahme.
4. Der Benutzer waehlt die Ziel-KW beziehungsweise die Ziel-Bestellwoche aus.
5. Die ausgewaehlten Profile werden mit Menge, Profil, Laenge und Projekt-/Kommissionsdaten in die Kindlimann Bestellliste uebertragen.

## Erweiterter Wunsch

Die Funktion soll nicht nur auf das aktuelle Kindlimann-Kontraktmaterial begrenzt sein.

Wichtig:

- Alle Profile aus der Profilstabzusammenfassung sollen grundsaetzlich uebertragbar sein.
- Kontraktmaterial bleibt als Stammdaten-/Dropdown-Liste relevant.
- Nicht-Kontraktmaterial soll trotzdem uebertragen werden koennen.
- Spaeter soll unterscheidbar sein, ob Profile als 6-Meter-Stangen oder als Zuschnitt bestellt werden sollen.

## Zuschnitt-Wunsch

Zusaetzlich soll untersucht werden, ob Profile, Rohre oder sonstige Materialien als Zuschnitt markiert werden koennen.

Moegliche Zielbedienung:

- Kontrollkaestchen fuer Bestellung als 6-Meter-Stange
- optionales Kontrollkaestchen oder Auswahl fuer Zuschnitt
- Uebernahme in eine passende Bestellstruktur
- spaetere Trennung zwischen Stangenmaterial und Zuschnittmaterial

## Nicht-Ziel

Dieser Arbeitsauftrag darf keine produktive Umsetzung ausfuehren.

Nicht erlaubt:

- keine Aenderung an der bestehenden Kindlimann Excel-Datei
- keine Aenderung am VBA-Code
- keine Aenderung an HiCAD-Systemdateien
- keine Aenderung an ReportManager-/Stuecklisten-Vorlagen
- keine produktive Integration in HiCAD
- keine automatische Ersetzung von HiCAD-Dateien

Dieser Auftrag ist eine Analyse und Konzeptklaerung.

## Zu untersuchende Quellen

1. Repository `kduerst/kindlimann-bestellvorlage`
   - aktuelle Bestelllistenstruktur
   - Konfiguration
   - PDF-/Mail-/Profil-Logik

2. Repository `kduerst/glndr-konfigurator`
   - bestehende HiCAD-nahe Automatisierung
   - Ablagepfade
   - Stuecklisten-/PartsList-Zielpfade
   - ReportManager-Hinweise
   - Output-Service und Pfad-Service

## Vorlaeufig bekannte Anknuepfungspunkte aus GLNDR

Im GLNDR-Konfigurator existiert bereits Logik fuer Automation-Ausgabeziele, darunter:

- Auftragsordner
- Planungsordner
- Kontrollplanordner
- Ausfuehrungsplaene
- Stuecklisten-Zielpfad

Es gibt Hinweise auf ReportManager-/Stuecklisten-Konfigurationen, unter anderem:

- `PartsListReportSettingsName`
- `PartsListReportSettingsFileName`
- `PartsListReportSettingsPath`
- `PartsListReportManagerTask`

Daraus folgt: Fuer eine spaetere Umsetzung sollte zuerst geklaert werden, ob die Profilstabzusammenfassung aus HiCAD als stabile Excel-/XLSX-/CSV-Ausgabe erzeugt oder direkt durch eine vorhandene HiCAD-Reportvorlage beeinflusst werden kann.

## Vorlaeufige Zielarchitektur

Die robusteste Richtung ist wahrscheinlich nicht, eine HiCAD-Systemdatei direkt unkontrolliert zu veraendern.

Empfohlener Ansatz fuer spaetere Umsetzung:

```text
HiCAD / ReportManager
→ Profilstabzusammenfassung als Datei
→ Import-/Uebernahme-Modul
→ Auswahl per Kontrollkaestchen
→ Ziel-KW waehlen
→ Eintrag in Kindlimann Bestellliste
→ PDF/Mail wie bisher
```

## Zu klaerende technische Fragen

- Wo liegt die aktuelle HiCAD-/ReportManager-Vorlage fuer die Profilstabzusammenfassung?
- Ist die Profilstabzusammenfassung eine Excel-Datei, eine ReportManager-Ausgabe oder eine aus HiCAD generierte Tabelle?
- Kann diese Datei gefahrlos angepasst werden?
- Wird sie projektbezogen neu erzeugt oder ist sie eine zentrale Vorlage?
- Welche Spalten sind stabil vorhanden?
- Wo stehen Kommissionsnummer und Projektname?
- Wie werden Profile eindeutig benannt?
- Wo steht die auf 6-Meter-Stangen gerechnete Menge?
- Wie werden Zuschnitte in der aktuellen Ausgabe abgebildet?
- Soll die Uebernahme direkt in die produktive Bestellliste schreiben oder zuerst in eine Import-Tabelle?
- Wie wird verhindert, dass doppelte Positionen mehrfach uebernommen werden?
- Wie wird dokumentiert, aus welchem Projekt eine uebernommene Position stammt?

## Pruefkriterien

Der Auftrag ist erfuellt, wenn dokumentiert ist:

- ob die Idee grundsaetzlich sinnvoll ist
- welche Datenquelle verwendet werden soll
- welche Risiken bestehen
- welche Dateien zuerst gefunden werden muessen
- ob Excel/VBA, PowerShell, Codex oder HiCAD-Anpassung der richtige Umsetzungsweg ist
- welche Schritte vor einer produktiven Umsetzung notwendig sind

## Erwartetes Ergebnisdokument

Nach Abschluss ist folgendes Dokument zu erstellen oder zu aktualisieren:

```text
Dokumentation/Entscheidungen/Arbeitsauftrag_000_HiCAD_Profilstabuebernahme_Machbarkeit.md
```

Dieses Ergebnisdokument muss enthalten:

- Ziel
- Quellenlage
- gefundene Anknuepfungspunkte
- technische Einschaetzung
- Risiken
- empfohlene Zielarchitektur
- offene Fragen
- Entscheidung
- Empfehlung fuer den naechsten Auftrag

## Freigabe

Die finale Freigabe erfolgt durch Kevin.

Eine produktive Umsetzung darf erst nach separatem Folgeauftrag erfolgen.

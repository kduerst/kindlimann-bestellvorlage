# Arbeitsauftrag 001 - Analyse HiCAD Stueckliste Profilstab Export

Datum: 2026-06-25  
Status: Analyse abgeschlossen, keine funktionalen Aenderungen umgesetzt  
Arbeitsmodus: Analyse, Kopie und Dokumentation

## Ausgangslage

Issue #1 im Repository `kduerst/kindlimann-bestellvorlage` konnte lokal nicht direkt ueber den GitHub-Connector gelesen werden, weil der Authentifizierungstoken abgelaufen war. Die Analyse wurde lokal anhand der vorhandenen GLNDR-/HiCAD-Dateien fortgefuehrt.

Gesucht war sinngemaess die Excel-Liste mit dem Namen `HiCAD Stahlbau erweiterte Winkel 2023`.

## Gefundene HiCAD-Ausgangsdateien

Im GLNDR-Code wurde der relevante HiCAD-Report eindeutig gefunden.

Quelle im Code:

```text
C:\Meine Programme\GelaenderPlugin\Plugin\GLNDR.Konfigurator\Services\GLNDRAutomationOutputService.cs
```

Relevante Konstanten:

```text
PartsListReportSettingsName = "HiCAD_Stahlbau_erweiterte_Winkel_2023"
PartsListReportSettingsFileName = "HiCAD_Stahlbau_erweiterte_Winkel_2023.rm_settings"
PartsListReportSettingsPath = @"C:\HiCAD\sys\HiCAD_Stahlbau_erweiterte_Winkel_2023.rm_settings"
```

Gefundene Report-Settings:

```text
C:\HiCAD\sys\HiCAD_Stahlbau_erweiterte_Winkel_2023.rm_settings
```

Metadaten:

- Typ: `.rm_settings`
- Groesse: 125761 Byte
- Geaendert: 2025-11-24 12:05:06

Die `.rm_settings` verweist auf folgende Excel-Vorlage:

```text
C:\HiCAD\sys\HiCAD_Stahlbau.DE.2802.0_CH_erweiterte Winkel_Metallraum.xlsx
```

Metadaten der Excel-Ausgangsdatei:

- Typ: `.xlsx`
- Groesse: 506148 Byte
- Geaendert: 2023-07-18 15:23:30
- Makros: keine VBA-Projektdatei gefunden (`xl/vbaProject.bin` fehlt)

Hinweis: Es existiert zusaetzlich eine neuere aehnliche Datei:

```text
C:\HiCAD\sys\HiCAD_Stahlbau.DE.3001.1_CH_erweiterte Winkel.xlsx
```

Diese ist neuer, wird aber von der gefundenen `HiCAD_Stahlbau_erweiterte_Winkel_2023.rm_settings` nicht referenziert. Fuer Arbeitsauftrag 001 wurde deshalb die von GLNDR/HiCAD referenzierte 2023/2802-Vorlage als Ausgangsdatei dokumentiert.

## Erstellte Arbeitskopie

Da die Ausgangsdatei keine Makros enthaelt, wurde die Kopie als `.xlsx` erstellt.

Kopie:

```text
C:\Meine Programme\Kindlimann\HiCAD_Stueckliste_Profilstab_Export_2026.xlsx
```

Quelle:

```text
C:\HiCAD\sys\HiCAD_Stahlbau.DE.2802.0_CH_erweiterte Winkel_Metallraum.xlsx
```

Metadaten der Kopie:

- Typ: `.xlsx`
- Groesse: 506148 Byte
- Geaendert: 2023-07-18 15:23:30
- Inhaltliche Aenderungen: keine

## Vergleich mit Kindlimann-Bestellliste

Referenzdatei:

```text
C:\Meine Programme\Kindlimann\Kindlimann Bestellliste.xlsm
```

| Punkt | HiCAD-Kopie | Kindlimann-Bestellliste |
| --- | --- | --- |
| Dateityp | `.xlsx` | `.xlsm` |
| Makros | nein | ja |
| Zweck | HiCAD-Report-/Exportvorlage | Bestellformular mit Makros |
| Blattstruktur | viele fachliche HiCAD-Reportblaetter | Anleitung, Stammdaten, KW01-KW53 |
| Schutz | relevante HiCAD-Blaetter ungeschuetzt | KW-Blaetter geschuetzt |
| Buttons | keine Buttons auf `Profilstab - Zusammenfassung` | Formularbuttons auf KW-Blaettern |
| Drucklogik | Excel-Vorlage mit vorhandenen Seiteneinstellungen | VBA setzt Druckbereich und PDF-Export |
| Relevanter Tab | `Profilstab - Zusammenfassung` | KW-Blatt, z. B. `KW27` |

Wichtige strukturelle Differenz: Die HiCAD-Datei ist eine Reportvorlage mit Platzhaltern wie `!header!...` und `!table!...`. Die Kindlimann-Datei ist ein bedienbares Formular mit Datenvalidierung, Button-Makros, Schutzlogik und PDF-/Mail-Automation.

## Analyse Kindlimann-Menue-/Buttonbereich

Analysiertes Blatt: `KW27`

Die Buttons liegen rechts neben der Druck-/Tabellenflaeche und werden nicht gedruckt (`PrintObject = False`).

### Hauptbutton

Text:

```text
Ablegen und E-Mail vorbereiten
```

Makro:

```text
ExportAndMailKindlimannOrder
```

Form:

- Shape-Name: `Rounded Rectangle 7`
- ShapeType: `1`
- AutoShapeType: `5` (abgerundetes Rechteck)
- Placement: `3`
- Locked: `True`
- PrintObject: `False`

Position:

- TopLeftCell: `F10`
- BottomRightCell: `H13`
- Links: 842.25 pt / 297.13 mm
- Oben: 252.75 pt / 89.16 mm
- Breite: 206.00 pt / 72.67 mm
- Hoehe: 44.00 pt / 15.52 mm

Farben:

- Fuellung: `#1C1C1C`
- Rahmen: `#1C1C1C`
- Schriftfarbe: `#FFFFFF`

Rahmen:

- Linienstaerke: 1 pt

Schrift:

- Font: `Aptos`
- Groesse: 9 pt
- Fett: ja
- Horizontal zentriert: ja
- Vertikal zentriert: ja

### Weitere Buttons im gleichen Bereich

| Text | Makro | Position | Groesse | Fuellung | Rahmen | Schrift |
| --- | --- | --- | --- | --- | --- | --- |
| `+ Auftrag` | `AddOrderBlockButton` | F5-G5 | 96 x 25 pt | `#F4F5F6` | `#C4C7CC`, 1 pt | Aptos 9, fett, `#1C1C1C` |
| `- Auftrag` | `DeleteOrderBlock` | G5-H5 | 96 x 25 pt | `#FFFFFF` | `#C4C7CC`, 1 pt | Aptos 9, fett, `#56585C` |
| `- Profil` | `DeleteProfileRow` | F7-G8 | 96 x 25 pt | `#FFFFFF` | `#C4C7CC`, 1 pt | Aptos 9, fett, `#56585C` |
| `+` | `AddProfileRowForOrderButton` | D5 | 24 x 20 pt | `#FFFFFF` | `#C4C7CC`, 1 pt | Aptos 13, fett, `#1C1C1C` |

Abstand zur Tabelle:

- Die Hauptaktionsbuttons beginnen bei 842.25 pt.
- Der Tabellenbereich endet bei Spalte D.
- Der zuletzt umgesetzte Abstand betraegt 24 pt rechts der Tabelle.

## Pruefung Tab `Profilstab - Zusammenfassung`

Der im Auftrag genannte Tab `Profilstab Zusammenfassung` wurde nicht exakt gefunden. Gefunden wurde der fachlich passende Tab:

```text
Profilstab - Zusammenfassung
```

Eigenschaften:

- Sichtbar: ja
- UsedRange: `A1:K9`
- Zeilen: 9
- Spalten: 11
- Geschuetzt: nein
- Shapes: 0
- Seitenorientierung: Querformat
- Druckbereich: leer/nicht explizit gesetzt
- Seitenraender:
  - Links: 51.02 pt
  - Rechts: 51.02 pt
  - Oben: 56.69 pt
  - Unten: 56.69 pt

Inhaltliche Struktur:

| Bereich | Inhalt |
| --- | --- |
| A1 | Titel `Profilstab - Zusammenfassung` |
| G1/I1 | Kommission und Header-Platzhalter |
| A3-C5 | Objekt, Bauherr, Architekt mit Header-Platzhaltern |
| E3-F5 | Haupt-Bezeichnung, Ersteller, Erstellt am |
| A8-K8 | Tabellenkopf |
| A9-K9 | Tabellenplatzhalter fuer HiCAD-Export |

Tabellenkopf in Zeile 8:

- `Nr.`
- `Anzahl`
- `Bezeichnung`
- `Laenge (mm)`
- `Material`
- `Typ`
- `Schnittbreite`
- `Schnitt-Zuschlag`
- `Abst. Anfang`
- `Abst.Ende`

Formatbeispiele:

| Zelle | Text | Schrift | Groesse | Fett | Schriftfarbe | Fuellung |
| --- | --- | --- | --- | --- | --- | --- |
| A1 | Profilstab - Zusammenfassung | Arial | 20 | ja | `#0000A1` | `#FFFFFF` |
| G1 | Kommission: | Arial | 20 | ja | `#0000A1` | `#FFFFFF` |
| A8 | Nr. | Arial | 10 | ja | `#000000` | `#DCE6F1` |
| B8 | Anzahl | Arial | 10 | ja | `#000000` | `#DCE6F1` |
| C8 | Bezeichnung | Arial | 10 | ja | `#000000` | `#DCE6F1` |
| A9 | `!table!=ROW()-8` | Arial Narrow | 10 | nein | `#000000` | `#FFFFFF` |

Spaltenbreiten:

| Spalte | Breite |
| --- | --- |
| A | 7 |
| B | 8 |
| C | 36 |
| D | 13.14 |
| E | 20 |
| F | 29 |
| G | 12 |
| H | 15 |
| I | 13 |
| J | 13 |
| K | 0 |

Zeilenhoehen:

| Zeile | Hoehe |
| --- | --- |
| 1 | 40 |
| 2 | 15 |
| 3 | 15.5 |
| 4 | 15.5 |
| 5 | 15.5 |
| 6 | 15.5 |
| 7 | 15 |
| 8 | 15 |
| 9 | 15 |

## Befund

Die gefundene HiCAD-Quelle ist keine Makrodatei, sondern eine `.xlsx`-Reportvorlage, die ueber die HiCAD-Report-Settings `HiCAD_Stahlbau_erweiterte_Winkel_2023.rm_settings` referenziert wird.

Der relevante Profilstab-Tab ist bereits vorhanden, aber nicht im Bedien-/Formularstil der Kindlimann-Bestellliste. Er ist eine technische HiCAD-Exportvorlage mit Platzhaltern.

Der Kindlimann-Hauptbutton ist gut als visuelle und technische Referenz fuer spaetere Bedienbuttons geeignet:

- dunkler Primaerbutton
- weisse Schrift
- Aptos 9 fett
- 206 x 44 pt
- rechts neben der Tabelle
- nicht druckbar
- Makrobindung ueber `OnAction`

## Keine funktionalen Aenderungen

Es wurden keine funktionalen Aenderungen an der HiCAD-Vorlage, der Kopie oder der Kindlimann-Bestellliste vorgenommen.

Erstellt wurden nur:

- eine unveraenderte Arbeitskopie der HiCAD-Excel-Vorlage
- dieses Entscheidungs-/Analysedokument

## Offene Entscheidungen vor Implementierung

1. Umsetzung auf der von GLNDR referenzierten 2023/2802-Vorlage?
   - Empfehlung: Ja, weil `HiCAD_Stahlbau_erweiterte_Winkel_2023.rm_settings` exakt auf diese Datei zeigt.

2. Tabname `Profilstab - Zusammenfassung` beibehalten?
   - Empfehlung: Ja, solange HiCAD diesen Namen erwartet.

3. Kindlimann-Hauptbutton-Stil uebernehmen?
   - Empfehlung: Ja, aber nur ausserhalb des HiCAD-Exportbereichs und nicht druckbar.

4. Datei `.xlsx` bleiben lassen?
   - Empfehlung: Ja, solange keine Makros benoetigt werden.

5. Spaetere Implementierung nur Layout/Bedienung oder auch Exportlogik/Makros?
   - Aktueller Analysebefund: Fuer reine HiCAD-Reportvorlage sind keine Makros noetig.

## Empfehlung fuer den naechsten Auftrag

Arbeitsauftrag 002 muss vor Umsetzung angepasst werden:

- Tabname exakt auf `Profilstab - Zusammenfassung` setzen.
- Ausgangsvorlage exakt auf `C:\HiCAD\sys\HiCAD_Stahlbau.DE.2802.0_CH_erweiterte Winkel_Metallraum.xlsx` setzen.
- Neue Datei zunaechst als `.xlsx` fuehren, solange keine Makrofunktion eingebaut wird.
- Uebertragungslogik nicht in die HiCAD-Reportvorlage einbauen, solange unklar ist, ob HiCAD die Vorlage mit Makros akzeptiert.
- Button-Stil kann als reine Form/Shape-Vorlage uebernommen werden, aber ohne Makrobindung, falls `.xlsx` beibehalten wird.
- Fuer die eigentliche Uebertragung ist eine separate Makro-/Plugin-/Importer-Loesung zu pruefen.

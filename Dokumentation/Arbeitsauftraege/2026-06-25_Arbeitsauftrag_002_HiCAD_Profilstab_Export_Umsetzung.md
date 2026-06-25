# Arbeitsauftrag 002 - HiCAD-Profilstab-Export umsetzen

## Ziel

Nach Abschluss der Analyse aus Arbeitsauftrag 001 soll nun die erste effektive Umsetzung vorbereitet und ausgefuehrt werden.

Ziel ist eine neue HiCAD-Arbeitsdatei, die auf der bestehenden HiCAD-Stueckliste basiert und im Tab `Profilstab Zusammenfassung` eine Auswahl und Uebertragungsfunktion fuer die Kindlimann-Bestellliste erhaelt.

## Bezug zu Arbeitsauftrag 001

Grundlage ist:

```text
Dokumentation/Arbeitsauftraege/2026-06-25_Arbeitsauftrag_001_HiCAD_Stueckliste_Profilstab_Import.md
```

Vor Umsetzung muessen die von Codex abgelegten Ergebnisse aus Arbeitsauftrag 001 gesucht und gelesen werden.

Codex soll insbesondere nach folgenden Ablageorten suchen:

```text
Dokumentation/Entscheidungen/
Dokumentation/Arbeitsauftraege/
GitHub Issue #1
lokale Codex-Ausgaben im Arbeitsordner
```

Wenn keine Ergebnisdokumentation aus Arbeitsauftrag 001 gefunden wird, muss Codex dies dokumentieren und vor riskanten Aenderungen stoppen.

## Ausgangsdateien

### Originaldatei HiCAD

Die bestehende HiCAD-Datei liegt voraussichtlich unter:

```text
\\int.metallraum.ch\root\folderredirections\kedu\Desktop\Dateien Privat\HICAD\HICAD Script\Metallraum Programme
```

Moegliche Ausgangsnamen:

```text
HiCAD Stahlbau erweiterte Winkel 2022
HiCAD_Stahlbau_erweiterte_Winkel_2023
```

### Zieldatei

Die neue Arbeitsdatei heisst verbindlich:

```text
HiCAD_Stueckliste_Profilstab_Export_2026.xlsm
```

Falls keine Makros benoetigt werden, darf `.xlsx` verwendet werden. Sobald ein Button mit aktiver Funktion notwendig ist, ist `.xlsm` zu bevorzugen.

## Unveraenderliche Regel

Die Originaldatei darf nicht veraendert werden.

Die neue Datei muss zu Beginn eine Kopie der Originaldatei sein.

In der neuen Datei duerfen nur die hier definierten Stellen angepasst werden:

1. Tab `Profilstab Zusammenfassung`
2. neue Spalte ganz rechts
3. Kontrollkaestchen bzw. Auswahlzellen in dieser neuen Spalte
4. Button fuer die Uebertragung
5. minimal notwendiger VBA-/Makrocode fuer die Uebertragungsfunktion

Keine sonstigen Layout-, Formel-, Druckbereich-, Farb-, Schrift- oder Strukturveraenderungen.

## Umsetzung Teil A - Auswahlspalte

Im Tab `Profilstab Zusammenfassung` wird ganz rechts hinter der bestehenden Tabelle eine neue Spalte ergaenzt.

Ueberschrift:

```text
Uebertrag Kindlimann Bestellliste
```

Anforderungen:

- Spaltenbreite so setzen, dass die gesamte Ueberschrift sichtbar ist.
- Darunter je Profilstab eine Auswahlmoeglichkeit einrichten.
- Bevorzugt: echte Kontrollkaestchen, wenn robust.
- Alternative: Auswahlspalte mit WAHR/FALSCH oder Ja/Nein, wenn technisch stabiler.
- Alle Profilstaebe muessen auswaehlbar sein.
- Es duerfen nicht nur die Standardprofile der Kindlimann-Liste beruecksichtigt werden.

## Umsetzung Teil B - Button

Im Tab `Profilstab Zusammenfassung` wird ein Button ergaenzt.

Beschriftung:

```text
Auswahl in Bestellliste uebertragen
```

Button-Stil:

Der Button muss optisch dem Button entsprechen, der in der Kindlimann-Bestellliste das Menuefenster oeffnet.

Codex muss:

1. Die Kindlimann-Bestellliste oeffnen bzw. analysieren.
2. Den Menuefenster-Button identifizieren.
3. Groesse, Farbe, Schrift, Rahmen und Stil dokumentieren.
4. Den neuen Button in der HiCAD-Datei nach diesem Stil nachbilden.

## Umsetzung Teil C - Kalenderwoche

Vor der Uebertragung muss eine Ziel-Kalenderwoche gewaehlt werden koennen.

Moegliche robuste Varianten:

1. Eingabefeld / Dialog beim Klick auf den Button.
2. Dropdown-Zelle im Tab `Profilstab Zusammenfassung`.
3. Kleines Menuefenster analog zur Kindlimann-Bestellliste.

Codex soll die einfachste stabile Variante waehlen und dokumentieren.

## Umsetzung Teil D - Uebertragung in die Kindlimann-Bestellliste

Die markierten Profilstaebe sollen in den passenden KW-Tab der Kindlimann-Bestellliste uebertragen werden.

Beispiel:

```text
Auswahl KW 32
→ Einfuegen in Tab "KW 32"
```

Die Daten werden in die naechsten freien Zeilen geschrieben.

Zu uebernehmende Felder:

| Ziel in Kindlimann | Quelle aus HiCAD-Profilstab-Zusammenfassung |
| --- | --- |
| Auftrag / Projekt | falls vorhanden, sonst leer oder Benutzereingabe |
| Menge | Anzahl / Stueckzahl |
| Profil | Profilname / Profilbezeichnung |
| Laenge | Laenge / Stablaenge |
| Einheit | Stk. |
| Bemerkung | optional leer |

## Profilregel

Alle Profile aus der HiCAD-Liste muessen uebertragbar sein.

Die Uebertragung darf nicht auf diese fuenf Standardprofile beschraenkt werden:

- 40 x 15 x 2 mm
- FL 30 x 6 mm
- FL 40 x 10 mm
- FL 40 x 15 mm
- RD 10 mm

Wenn ein Profil nicht in der Kindlimann-Dropdownliste existiert, soll der Profiltext trotzdem in die Bestellliste geschrieben werden koennen.

## Sicherheitsanforderungen

Codex muss vor jeder schreibenden Aenderung sicherstellen:

- Es wird nicht in die Originaldatei geschrieben.
- Es wird nicht ungeprueft in die produktive Kindlimann-Datei geschrieben.
- Zuerst wird eine Testkopie verwendet.
- Bestehende Formatierung wird nicht unbeabsichtigt veraendert.
- Alle erzeugten Dateien und Pfade werden dokumentiert.

## Erwartetes Vorgehen

1. Ergebnisse aus Arbeitsauftrag 001 suchen und lesen.
2. Gefundene Ergebnisse kurz zusammenfassen.
3. Original-HiCAD-Datei finden.
4. Kopie `HiCAD_Stueckliste_Profilstab_Export_2026.xlsm` erstellen.
5. Kindlimann-Bestellliste analysieren, speziell den Menuefenster-Button.
6. Tab `Profilstab Zusammenfassung` pruefen.
7. Neue Auswahlspalte ergaenzen.
8. Kontrollkaestchen/Auswahlzellen ergaenzen.
9. Button im Kindlimann-Stil ergaenzen.
10. Einfache KW-Auswahl vorsehen.
11. Uebertragungslogik zuerst mit Testdatei testen.
12. Ergebnis dokumentieren.

## Ergebnisdokumentation

Nach Abschluss muss ein Ergebnisdokument erstellt werden:

```text
Dokumentation/Entscheidungen/Arbeitsauftrag_002_HiCAD_Profilstab_Export_Umsetzung.md
```

Dieses Dokument muss enthalten:

- Ziel
- gefundene Ausgangsdatei
- erstellte Kopie
- geaenderte Stellen
- Nachweis, dass die Originaldatei unveraendert blieb
- Button-Stil der Kindlimann-Bestellliste
- technische Umsetzung
- Testresultat
- Blockaden
- Empfehlung fuer den naechsten Auftrag

## Nicht Teil dieses Auftrags

Noch nicht umsetzen:

- Zuschnittliste
- Anschnittwinkel
- Gradangaben
- automatische Optimierung von Stablaengen
- Gruppierung mehrerer Zuschnitte
- direkte produktive Uebertragung ohne Testlauf

Diese Punkte gehoeren in einen spaeteren Arbeitsauftrag 003.

## Akzeptanzkriterien

Der Auftrag ist erfolgreich, wenn:

- die neue Datei `HiCAD_Stueckliste_Profilstab_Export_2026.xlsm` bzw. `.xlsx` erstellt wurde,
- die Originaldatei unveraendert blieb,
- im Tab `Profilstab Zusammenfassung` rechts eine Auswahlspalte vorhanden ist,
- alle Profilstaebe auswaehlbar sind,
- der Button optisch dem Menuefenster-Button der Kindlimann-Bestellliste entspricht,
- eine KW ausgewaehlt werden kann,
- ausgewaehlte Profile testweise in den passenden KW-Tab der Kindlimann-Bestellliste uebertragen werden koennen,
- die Umsetzung dokumentiert ist.

# Arbeitsauftrag 002 - HiCAD-Profilstab-Export direkt mit Bestellliste verknuepfen

## Ziel

Es soll praktisch geprueft und so weit wie moeglich direkt umgesetzt werden, ob die HiCAD-Profilstab-Liste und die Kindlimann-Bestellliste funktional miteinander verknuepft werden koennen.

Ziel ist eine erste funktionierende Loesung:

```text
HiCAD_Stueckliste_Profilstab_Export_2026
→ Profilstab - Zusammenfassung
→ ausgewaehlte Profilstaebe
→ Kindlimann Bestellliste
→ passender KW-Tab
```

Die Struktur aus Arbeitsauftrag 001 passt. Es sollen keine grossen Umbauten erfolgen, sondern nur die benoetigten Punkte ergaenzt werden.

## Grundlage

Vor Umsetzung muss Codex das Ergebnis aus Arbeitsauftrag 001 beruecksichtigen:

```text
Dokumentation/Entscheidungen/Arbeitsauftrag_001_HiCAD_Stueckliste_Profilstab_Export_Analyse.md
```

Wichtige Erkenntnisse daraus:

- Die relevante HiCAD-Ausgangsvorlage ist:

```text
C:\HiCAD\sys\HiCAD_Stahlbau.DE.2802.0_CH_erweiterte Winkel_Metallraum.xlsx
```

- Die Report-Settings verweisen auf:

```text
C:\HiCAD\sys\HiCAD_Stahlbau_erweiterte_Winkel_2023.rm_settings
```

- Der exakte relevante Blattname lautet:

```text
Profilstab - Zusammenfassung
```

- Die Ausgangsdatei ist eine `.xlsx` ohne Makros.
- Die Kindlimann-Bestellliste ist eine `.xlsm` mit Makros.
- Der Menue-/Hauptbutton der Kindlimann-Bestellliste ist optische Referenz.

## Zentrale Frage dieses Auftrags

Kann die HiCAD-Profilstab-Liste praktikabel mit der Kindlimann-Bestellliste verknuepft werden?

Codex soll diese Frage nicht nur theoretisch beantworten, sondern eine minimale, testbare Umsetzung erstellen, sofern technisch moeglich.

## Ziel-Dateien

### HiCAD-Arbeitsdatei

Arbeitsdatei:

```text
C:\Meine Programme\Kindlimann\HiCAD_Stueckliste_Profilstab_Export_2026.xlsx
```

Falls fuer die direkte Button-Funktion Makros in dieser Datei zwingend erforderlich sind und HiCAD diese Datei danach noch akzeptiert, darf eine `.xlsm`-Variante als Testdatei erstellt werden:

```text
C:\Meine Programme\Kindlimann\HiCAD_Stueckliste_Profilstab_Export_2026.xlsm
```

Die Originaldatei unter `C:\HiCAD\sys\...` darf nicht veraendert werden.

### Kindlimann-Testdatei

Die Uebertragung darf zuerst nicht ungeprueft in die produktive Kindlimann-Datei erfolgen.

Codex soll fuer den ersten Test eine Kopie verwenden, z. B.:

```text
C:\Meine Programme\Kindlimann\Kindlimann Bestellliste TEST.xlsm
```

Erst wenn der Test funktioniert, kann Kevin entscheiden, ob die Logik produktiv uebernommen wird.

## Unveraenderliche Regel

Die HiCAD-Vorlage soll optisch und strukturell praktisch unveraendert bleiben.

Erlaubt sind nur diese Aenderungen im Tab `Profilstab - Zusammenfassung`:

1. neue Auswahlspalte ganz rechts
2. Kontrollkaestchen bzw. WAHR/FALSCH-Zellen in dieser Spalte
3. KW-Auswahl oder Dialog
4. Button fuer `Auswahl in Bestellliste uebertragen`
5. minimal notwendige Logik fuer die Uebertragung

Keine sonstigen Format-, Formel-, Spalten-, Druckbereich-, Farb-, Schrift- oder Strukturveraenderungen.

## Umsetzung Teil A - Exakter Blattname

Der Auftrag verwendet ab sofort ausschliesslich den exakten Blattnamen:

```text
Profilstab - Zusammenfassung
```

Nicht verwenden:

```text
Profilstab Zusammenfassung
```

## Umsetzung Teil B - Auswahlspalte mit Kontrollkaestchen

Im Blatt `Profilstab - Zusammenfassung` wird rechts neben den bestehenden Daten eine neue Spalte ergaenzt.

Ueberschrift:

```text
Uebertrag Kindlimann Bestellliste
```

Anforderungen:

- Spaltenbreite so setzen, dass die gesamte Ueberschrift sichtbar ist.
- Die bestehende Formatierung darf sonst nicht veraendert werden.
- Pro Profilstab soll ein Kontrollkaestchen vorhanden sein.
- Das Kontrollkaestchen soll intern mit einer Zellverknuepfung arbeiten.
- Die verknuepfte Zelle soll `WAHR` oder `FALSCH` enthalten.
- Optisch soll die Bedienung einfach bleiben: Benutzer klickt das Kaestchen an.
- Alle Profilstaebe muessen auswaehlbar sein.
- Nicht nur die fuenf Standardprofile der Kindlimann-Liste duerfen uebertragen werden.

Falls echte Formular-Kontrollkaestchen im HiCAD-Reportkontext technisch instabil sind, soll Codex dies dokumentieren und eine optisch einfache Ja/Nein-Alternative bauen. Prioritaet hat aber das Kontrollkaestchen mit WAHR/FALSCH.

## Umsetzung Teil C - Button mit direkter Funktion

Im Blatt `Profilstab - Zusammenfassung` wird ein Button ergaenzt.

Beschriftung:

```text
Auswahl in Bestellliste uebertragen
```

Der Button soll direkt funktionieren und die markierten Daten in die Kindlimann-Bestellliste uebertragen.

Button-Stil:

Der Button muss optisch dem Hauptbutton der Kindlimann-Bestellliste entsprechen:

- Textreferenz: `Ablegen und E-Mail vorbereiten`
- dunkler Primaerbutton
- Fuellung: `#1C1C1C`
- Rahmen: `#1C1C1C`
- Schriftfarbe: `#FFFFFF`
- Schrift: Aptos, 9 pt, fett
- abgerundetes Rechteck
- Groesse ca. 206 x 44 pt
- nicht druckbar

Falls eine identische Uebernahme technisch nicht exakt moeglich ist, soll Codex den Stil so nahe wie moeglich nachbauen.

## Umsetzung Teil D - Kalenderwoche

Vor dem Uebertragen muss eine Kalenderwoche gewaehlt werden koennen.

Bevorzugte einfache Loesung:

- Beim Klick auf den Button erscheint eine Eingabeaufforderung fuer die KW.
- Eingabe z. B. `27` oder `KW27`.
- Codex normalisiert die Eingabe auf den Zieltab `KW27`.

Alternative:

- Dropdown-Zelle im Blatt `Profilstab - Zusammenfassung`, falls stabiler und klarer.

Ziel: so einfach wie moeglich, aber funktionierend.

## Umsetzung Teil E - Direkte Uebertragung

Beim Klick auf den Button sollen alle Zeilen mit gesetztem Kontrollkaestchen (`WAHR`) gelesen werden.

Aus `Profilstab - Zusammenfassung` sollen fuer Schritt 1 diese Werte uebertragen werden:

| Ziel in Kindlimann | Quelle aus HiCAD-Profilstab-Zusammenfassung |
| --- | --- |
| Auftrag / Projekt | falls nicht vorhanden: leer oder Eingabewert |
| Menge | Spalte `Anzahl` |
| Profil | Spalte `Bezeichnung` |
| Laenge | Spalte `Laenge (mm)` |
| Einheit | `Stk.` |
| Bemerkung | leer |

Ziel:

```text
Kindlimann Bestellliste TEST.xlsm
→ gewaehlter KW-Tab, z. B. KW27
→ naechste freie Bestellzeilen
```

Wenn ein Profil nicht in der Kindlimann-Dropdownliste existiert, soll der Profiltext trotzdem in die Zielzelle geschrieben werden. Die Dropdownliste darf die Uebertragung nicht blockieren.

## Funktionale Mindestanforderung

Am Ende dieses Arbeitsauftrags soll mindestens ein Test moeglich sein:

1. HiCAD-Testdatei oeffnen.
2. Im Blatt `Profilstab - Zusammenfassung` ein oder mehrere Profilstaebe markieren.
3. Button klicken.
4. KW auswaehlen oder eingeben.
5. Kindlimann-Testdatei wird geoeffnet oder verwendet.
6. Markierte Profilstaebe werden in den passenden KW-Tab geschrieben.
7. Benutzer erhaelt eine Rueckmeldung, wie viele Positionen uebertragen wurden.

## Wichtige technische Pruefung

Codex muss waehrend der Umsetzung klaeren und dokumentieren:

1. Funktioniert direkte Verknuepfung aus der HiCAD-Excel-Datei in die Kindlimann-Excel-Datei?
2. Muss die HiCAD-Datei dafuer als `.xlsm` gespeichert werden?
3. Akzeptiert HiCAD diese geaenderte Datei weiterhin als Reportvorlage?
4. Bleiben die Kontrollkaestchen nach einem HiCAD-Export erhalten?
5. Werden die neuen Spalten/Shapes durch HiCAD ueberschrieben oder bleiben sie stabil?

Wenn Punkt 3 bis 5 nicht sicher beantwortet werden koennen, soll Codex trotzdem eine Testumsetzung erstellen, aber klar dokumentieren, dass die Produktivtauglichkeit noch freigegeben werden muss.

## Sicherheitsanforderungen

- Original-HiCAD-Datei nicht veraendern.
- Produktive Kindlimann-Bestellliste nicht ungeprueft veraendern.
- Fuer den Test mit Kopien arbeiten.
- Nur definierte Stellen anpassen.
- Ergebnis dokumentieren.
- Falls etwas nicht stabil moeglich ist, stoppen und begruenden.

## Erwartetes Vorgehen

1. Ergebnis aus Arbeitsauftrag 001 lesen.
2. HiCAD-Arbeitskopie pruefen.
3. Kindlimann-Testkopie erstellen oder vorhandene Testkopie verwenden.
4. Blatt `Profilstab - Zusammenfassung` oeffnen.
5. Neue Spalte `Uebertrag Kindlimann Bestellliste` ergaenzen.
6. Kontrollkaestchen mit WAHR/FALSCH-Zellverknuepfung einfuegen.
7. Button im Kindlimann-Hauptbutton-Stil einfuegen.
8. Button-Funktion fuer direkte Uebertragung umsetzen.
9. Test mit mindestens zwei markierten Profilstaeben durchfuehren.
10. Ergebnis dokumentieren.

## Ergebnisdokumentation

Nach Abschluss muss ein Ergebnisdokument erstellt werden:

```text
Dokumentation/Entscheidungen/Arbeitsauftrag_002_HiCAD_Profilstab_Export_Umsetzung.md
```

Dieses Dokument muss enthalten:

- Ziel
- verwendete Ausgangsdateien
- erstellte Testdateien
- geaenderte Stellen
- Nachweis, dass die Originaldateien unveraendert blieben
- technische Umsetzung der Kontrollkaestchen
- technische Umsetzung des Buttons
- technische Umsetzung der KW-Auswahl
- technische Umsetzung der Uebertragung
- Testresultat
- bekannte Grenzen
- Antwort auf die Frage, ob die Dateien funktional verknuepfbar sind
- Empfehlung fuer produktive Uebernahme

## Nicht Teil dieses Auftrags

Noch nicht umsetzen:

- Zuschnittliste
- Anschnittwinkel
- Gradangaben
- automatische Optimierung von Stablaengen
- Gruppierung mehrerer Zuschnitte
- produktive Freigabe ohne Testauswertung

Diese Punkte gehoeren in einen spaeteren Arbeitsauftrag 003.

## Akzeptanzkriterien

Der Auftrag ist erfolgreich, wenn:

- die direkte Verknuepfbarkeit praktisch getestet wurde,
- die neue Auswahlspalte im Blatt `Profilstab - Zusammenfassung` vorhanden ist,
- Kontrollkaestchen oder eine dokumentierte stabile Alternative vorhanden sind,
- die Auswahl intern WAHR/FALSCH oder Ja/Nein abbildet,
- der Button optisch dem Kindlimann-Hauptbutton entspricht,
- der Button die Uebertragung direkt ausfuehrt,
- eine KW ausgewaehlt/eingegeben werden kann,
- ausgewaehlte Profilstaebe testweise in die Kindlimann-Testdatei uebertragen werden,
- alle Profile uebertragbar sind, auch wenn sie nicht in der Kindlimann-Dropdownliste stehen,
- die Originaldateien unveraendert blieben,
- die Umsetzung dokumentiert ist.

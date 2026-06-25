# Arbeitsauftrag 001 - HiCAD-Stueckliste Profilstab-Export in Bestellliste

## Ziel

Die bestehende Kindlimann-Bestellliste soll perspektivisch aus einer von HiCAD erzeugten Excel-Stueckliste befuellt werden koennen.

Ausgangspunkt ist die HiCAD-Stueckliste bzw. Excel-Liste mit dem bisherigen Namen:

```text
HiCAD_Stahlbau_erweiterte_Winkel_2023
```

bzw. eine bestehende Vorlage aus dem Umfeld:

```text
HiCAD Stahlbau erweiterte Winkel 2022
```

Die alte/bestehende HiCAD-Liste darf nicht direkt veraendert oder ueberschrieben werden. Es soll zwingend eine neue Kopie bzw. neue Vorlage entstehen, die technisch erweitert wird.

## Bekannter lokaler Suchpfad

Die bestehende HiCAD-Liste soll lokal bzw. im Netzwerkpfad gesucht werden unter:

```text
\\int.metallraum.ch\root\folderredirections\kedu\Desktop\Dateien Privat\HICAD\HICAD Script\Metallraum Programme
```

Codex soll zuerst in diesem Pfad nach der passenden Datei suchen und dokumentieren, welche Datei als Ausgangsdatei gefunden wurde.

## Neuer Dateiname

Die neue Arbeitsdatei soll nicht Kindlimann im Namen enthalten, weil sie allgemein fuer HiCAD-Stuecklisten-/Profilstab-/Zuschnitt-Auswertungen verwendbar bleiben soll.

Verbindlicher Arbeitsname fuer die neue Kopie:

```text
HiCAD_Stueckliste_Profilstab_Export_2026.xlsx
```

Falls aus technischen Gruenden eine Makro-Datei noetig ist, darf Codex alternativ diese Endung verwenden:

```text
HiCAD_Stueckliste_Profilstab_Export_2026.xlsm
```

Die Originaldatei bleibt unveraendert.

## Wichtige Grundregel

Keine direkte Veraenderung der bestehenden Originaldatei:

```text
HiCAD Stahlbau erweiterte Winkel 2022
HiCAD_Stahlbau_erweiterte_Winkel_2023
```

Stattdessen:

1. Originaldatei identifizieren.
2. Eine Kopie erstellen.
3. Die Kopie unter dem Namen `HiCAD_Stueckliste_Profilstab_Export_2026.xlsx` bzw. `.xlsm` weiterentwickeln.
4. Erst diese neue Datei fuer Tests und Anpassungen verwenden.
5. Die Originalformatierung der Kopie vollstaendig erhalten.

## Strikte Formatierungsregel

Die neue Datei soll beim Start absolut identisch zur vorhandenen HiCAD-Stueckliste bleiben.

Erlaubt ist nur diese gezielte Erweiterung:

- Im Tab `Profilstab Zusammenfassung` wird ganz rechts hinter den bestehenden Daten eine neue Spalte eingefuegt.
- Diese neue Spalte erhaelt die Ueberschrift:

```text
Uebertrag Kindlimann Bestellliste
```

- Die Spaltenbreite dieser neuen Spalte muss so gesetzt werden, dass die komplette Ueberschrift sichtbar ist.
- In den Zeilen darunter werden Kontrollkaestchen oder eine gleichwertige Auswahlmoeglichkeit eingefuegt.
- Alle bisherigen Spalten, Zeilen, Formate, Formeln, Druckbereiche, Farben, Schriftarten und Layouts muessen unveraendert bleiben.
- Keine sonstige Formatierung darf veraendert werden.
- Keine bestehenden Inhalte duerfen verschoben, geloescht oder umformatiert werden.

## Button-Regel

Unten im Tab `Profilstab Zusammenfassung` soll ein Button ergaenzt werden.

Button-Ziel:

```text
Auswahl in Bestellliste uebertragen
```

Anforderungen:

- Position: unten im Tab `Profilstab Zusammenfassung`, ausserhalb der bestehenden Datenstruktur, ohne bestehende Inhalte zu ueberdecken.
- Groesse, Farbe und Stil sollen moeglichst identisch zu den Buttons der bestehenden Kindlimann-Bestellliste sein.
- Ziel ist ein einheitliches Erscheinungsbild zwischen HiCAD-Exportdatei und Kindlimann-Bestellliste.
- Falls die exakte Button-Formatierung aus der Kindlimann-Datei technisch nicht automatisch ausgelesen werden kann, soll Codex sie zuerst dokumentieren und dann manuell in der neuen HiCAD-Datei nachbilden.

## Ausgangslage

Die Kindlimann-Bestellliste benoetigt nur wenige Werte:

- Auftrag / Projekt
- Menge / Stueckzahl
- Profilname
- Laenge
- Einheit
- Bemerkung optional
- Kalenderwoche fuer die Einfuegung

Die HiCAD-Stueckliste enthaelt deutlich mehr Informationen. Fuer diesen ersten Arbeitsauftrag sollen nur die Felder betrachtet werden, die fuer die Kindlimann-Bestellliste notwendig sind.

## Gewuenschter Ablauf in der HiCAD-Excel-Liste

In der neuen Kopie der HiCAD-Stueckliste soll ein Tab vorhanden oder erweitert werden:

```text
Profilstab Zusammenfassung
```

In diesem Tab soll hinter jedem Profilstab eine Auswahlmoeglichkeit vorhanden sein.

Gewuenschte Bedienlogik:

1. Benutzer erzeugt/exportiert die HiCAD-Stueckliste.
2. Benutzer oeffnet den Tab `Profilstab Zusammenfassung`.
3. Alle Profilstaebe werden sichtbar aufgelistet.
4. Hinter jedem Profil gibt es in der neuen Spalte ein Kontrollkaestchen oder eine gleichwertige Auswahlspalte.
5. Benutzer markiert die Profilstaebe, die bestellt werden sollen.
6. Benutzer waehlt eine Kalenderwoche aus.
7. Benutzer klickt den Button `Auswahl in Bestellliste uebertragen`.
8. Die ausgewaehlten Profilstaebe werden in die Kindlimann-Bestellliste uebertragen.

## Wichtig: Alle Profile muessen auswaehlbar sein

Die Auswahl darf nicht nur auf die bisher in der Kindlimann-Bestellliste hinterlegten Standardprofile beschraenkt sein.

Bisherige Kindlimann-Profile:

- 40 x 15 x 2 mm
- FL 30 x 6 mm
- FL 40 x 10 mm
- FL 40 x 15 mm
- RD 10 mm

Trotzdem muessen auch andere Profilstaebe aus der HiCAD-Liste auswaehlbar und uebertragbar sein.

Grund:

Die HiCAD-Stueckliste kann mehr Profile enthalten als die standardisierte Kindlimann-Bestellliste. Fuer Sonderfaelle muss trotzdem jedes benoetigte Profil in die Bestellung uebernommen werden koennen.

## Import-/Uebertragungsziel

Die Uebertragung soll in die bestehende Kindlimann-Bestellliste erfolgen.

Dort soll anhand der gewaehlten Kalenderwoche der passende KW-Tab ausgewaehlt werden.

Beispiel:

```text
Auswahl KW 32
→ Einfuegen in Tab "KW 32" der Kindlimann-Bestellliste
```

Die Daten sollen in die naechsten freien Zeilen geschrieben werden.

## Zu uebernehmende Daten fuer Schritt 1

Aus `Profilstab Zusammenfassung` sollen fuer den ersten Schritt nur die Daten uebernommen werden, die in der Kindlimann-Bestellliste benoetigt werden:

| Kindlimann-Feld | Quelle aus HiCAD-Liste | Hinweis |
| --- | --- | --- |
| Auftrag / Projekt | falls vorhanden, sonst leer oder Eingabefeld | Muss fachlich geprueft werden |
| Menge | Stueckzahl / Anzahl | Muss aus der Zusammenfassung kommen |
| Profil | Profilname / Bezeichnung | Muss auch fuer nicht standardisierte Profile moeglich sein |
| Laenge | Laenge / Einheitslaenge | Normalfall vermutlich Stabstange ca. 6100 mm |
| Einheit | Stk. oder definierte Einheit | Vorerst einfach halten |
| Bemerkung | optional | Kann leer bleiben |

## Normalfall

Normalerweise werden Profilstaebe als Stangen bestellt.

Annahme fuer den Normalfall:

```text
1 Stange ≈ 6100 mm
```

Die Kindlimann-Bestellliste soll daher im ersten Schritt vor allem Profilstab-Bestellungen aus der Zusammenfassung uebernehmen.

## Nicht Teil dieses ersten Arbeitsschritts

Der Zuschnittimport mit einzelnen Laengen, Anschnittwinkeln, Gradangaben und mehreren Schnittpositionen ist bewusst noch nicht Teil der Umsetzung.

Dieser Punkt wird als zweiter Arbeitsauftrag vorbereitet.

## Folgeidee / Arbeitsauftrag 002

Wenn der Import aus `Profilstab Zusammenfassung` funktioniert, soll spaeter ein zweiter Schritt spezifiziert werden:

```text
HiCAD-Zuschnittliste → Kindlimann-Bestellliste / Zuschnittbestellung
```

Moegliche Inhalte von Arbeitsauftrag 002:

- Auswahl einzelner Zuschnittpositionen
- Uebernahme von Profil
- Uebernahme von Laenge
- Uebernahme von Anschnittwinkel links/rechts
- Uebernahme von Gradangaben
- Unterscheidung zwischen ganzer Stange und zugeschnittenem Profil
- Pruefen, ob mehrere Zuschnitte sinnvoll gruppiert werden koennen

Beispiel:

```text
Profil A: 2000 mm + 3000 mm
→ entweder als ganze Stange bestellen
→ oder als definierte Zuschnittpositionen uebernehmen
```

## Technische Leitplanken fuer Codex

1. Zuerst analysieren, nicht sofort produktiv implementieren.
2. Keine bestehende HiCAD-Originaldatei ueberschreiben.
3. Keine bestehende Kindlimann-Produktivdatei unkontrolliert veraendern.
4. Keine bestehende Formatierung der HiCAD-Datei veraendern, ausser der neuen Auswahlspalte und dem neuen Button.
5. Falls Makros/VBA benoetigt werden, diese getrennt dokumentieren.
6. Aenderungen klein halten.
7. Importlogik so bauen, dass sie spaeter erweitert werden kann.
8. Zuerst mit einer Testkopie arbeiten.
9. Ergebnis dokumentieren unter `Dokumentation/Entscheidungen/`.

## Erwartetes Vorgehen fuer Codex

Codex soll schrittweise arbeiten:

1. Datei im angegebenen Pfad suchen.
2. Gefundene Ausgangsdatei dokumentieren.
3. Sicherheitskopie erstellen.
4. Neue Arbeitsdatei mit Namen `HiCAD_Stueckliste_Profilstab_Export_2026.xlsx` oder `.xlsm` erstellen.
5. Tab `Profilstab Zusammenfassung` finden und pruefen.
6. Bestehende Spalten, Formeln, Formatierung und Druckbereich protokollieren.
7. Rechts neben den bestehenden Daten eine neue Spalte `Uebertrag Kindlimann Bestellliste` ergaenzen.
8. Kontrollkaestchen/Auswahl in dieser Spalte einfuegen.
9. Button im Stil der Kindlimann-Bestellliste ergaenzen.
10. Technisches Konzept fuer die Uebertragung in die Kindlimann-Bestellliste dokumentieren.
11. Erst danach minimale Testumsetzung starten.

## Erwartetes Ergebnis dieses Arbeitsauftrags

Codex soll zuerst liefern:

1. Analyse der vorhandenen Dateien und Spalten.
2. Bestaetigung der neuen Kopie und des Dateinamens.
3. Nachweis, dass die Originaldatei unveraendert bleibt.
4. Nachweis, dass keine bestehende Formatierung veraendert wurde.
5. Vorschlag fuer die technische Uebertragung:
   - VBA
   - Power Query
   - Office Script
   - manuelle Export-/Importdatei
   - oder anderes robustes Verfahren
6. Entscheidungsvorlage, welche Variante fuer Metallraum am stabilsten ist.
7. Danach erst eine minimale Testumsetzung.

## Akzeptanzkriterien

Der Arbeitsauftrag gilt als fachlich erfolgreich vorbereitet, wenn:

- eine neue Kopie der HiCAD-Stueckliste erstellt oder vorgesehen ist,
- die Originaldatei unveraendert bleibt,
- der Name `HiCAD_Stueckliste_Profilstab_Export_2026` verwendet wird,
- der Tab `Profilstab Zusammenfassung` als Quelle definiert ist,
- nur dort eine neue Auswahlspalte ergaenzt wird,
- die neue Spalte `Uebertrag Kindlimann Bestellliste` heisst,
- die Spaltenbreite so gesetzt ist, dass die ganze Ueberschrift sichtbar ist,
- alle Profilstaebe auswaehlbar sind,
- nicht nur die fuenf Kindlimann-Standardprofile uebertragen werden koennen,
- die Ziel-KW auswaehlbar ist,
- der Button optisch zur Kindlimann-Bestellliste passt,
- die Datenstruktur fuer die Kindlimann-Bestellliste klar dokumentiert ist,
- der Zuschnittimport als separater Folgeauftrag abgegrenzt ist.

## Offene Fragen an Kevin

Vor Implementierung klaeren:

1. Heisst der Tab exakt `Profilstab Zusammenfassung` oder leicht anders?
2. Welche Spaltennamen sind in diesem Tab vorhanden?
3. Gibt es bereits eine Spalte fuer Projekt/Auftrag?
4. Soll der Transfer direkt in die Kindlimann-XLSM erfolgen oder zuerst ueber eine Zwischen-Datei?
5. Welche Excel-Version wird verwendet?
6. Darf VBA verwendet werden oder soll moeglichst ohne Makros gearbeitet werden?
7. Wo genau liegt die produktive Kindlimann-Bestellliste lokal beim Benutzer?

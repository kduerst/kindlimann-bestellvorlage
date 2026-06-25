# Arbeitsauftrag 001 - HiCAD-Stueckliste Profilstab-Import in Bestellliste

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

Die alte/bestehende HiCAD-Liste darf nicht direkt veraendert oder ueberschrieben werden. Es soll eine neue Kopie bzw. neue Vorlage entstehen, die technisch erweitert wird.

## Wichtige Grundregel

Keine direkte Veraenderung der bestehenden Originaldatei:

```text
HiCAD Stahlbau erweiterte Winkel 2022
```

Stattdessen:

1. Originaldatei identifizieren.
2. Eine Kopie erstellen.
3. Die Kopie unter einem neuen, neutralen Namen weiterentwickeln.
4. Erst diese neue Datei fuer weitere Tests und Anpassungen verwenden.

Der neue Name soll nicht auf Kindlimann festgelegt sein, weil die Liste allgemein fuer HiCAD-/Stuecklisten-/Zuschnitt-Zwecke nutzbar bleiben soll.

Namensvorschlaege:

```text
HiCAD_Profilstab_Stueckliste_2026.xlsx
HiCAD_Stahlbau_Profilstab_Export_2026.xlsx
HiCAD_Profilstab_Zuschnitt_Export_2026.xlsx
```

Codex soll einen passenden Namen vorschlagen und vor Umsetzung dokumentieren.

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
4. Hinter jedem Profil gibt es ein Kontrollkaestchen oder eine vergleichbare Auswahlspalte.
5. Benutzer markiert die Profilstaebe, die bestellt werden sollen.
6. Benutzer waehlt eine Kalenderwoche aus.
7. Die ausgewaehlten Profilstaebe werden in die Kindlimann-Bestellliste uebertragen.

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
| Auftrag / Projekt | falls vorhanden, sonst leer oder Eingabefeld | Muss spaeter fachlich geklaert werden |
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
4. Falls Makros/VBA benoetigt werden, diese getrennt dokumentieren.
5. Aenderungen klein halten.
6. Importlogik so bauen, dass sie spaeter erweitert werden kann.
7. Zuerst mit einer Testkopie arbeiten.
8. Ergebnis dokumentieren unter `Dokumentation/Entscheidungen/`.

## Erwartetes Ergebnis dieses Arbeitsauftrags

Codex soll zuerst liefern:

1. Analyse der vorhandenen Dateien und Spalten.
2. Vorschlag fuer den neuen Namen der HiCAD-Kopie.
3. Vorschlag fuer die technische Uebertragung:
   - VBA
   - Power Query
   - Office Script
   - manuelle Export-/Importdatei
   - oder anderes robustes Verfahren
4. Entscheidungsvorlage, welche Variante fuer Metallraum am stabilsten ist.
5. Danach erst eine minimale Testumsetzung.

## Akzeptanzkriterien

Der Arbeitsauftrag gilt als fachlich erfolgreich vorbereitet, wenn:

- eine neue Kopie der HiCAD-Stueckliste vorgesehen ist,
- die Originaldatei unveraendert bleibt,
- der Tab `Profilstab Zusammenfassung` als Quelle definiert ist,
- alle Profilstaebe auswaehlbar sind,
- nicht nur die fuenf Kindlimann-Standardprofile uebertragen werden koennen,
- die Ziel-KW auswaehlbar ist,
- die Datenstruktur fuer die Kindlimann-Bestellliste klar dokumentiert ist,
- der Zuschnittimport als separater Folgeauftrag abgegrenzt ist.

## Offene Fragen an Kevin

Vor Implementierung klaeren:

1. Wo liegt die aktuelle HiCAD-Excel-Vorlage lokal oder im Repository?
2. Wie heisst der Tab exakt: `Profilstab Zusammenfassung` oder anders?
3. Welche Spaltennamen sind in diesem Tab vorhanden?
4. Gibt es bereits eine Spalte fuer Projekt/Auftrag?
5. Soll der Transfer direkt in die Kindlimann-XLSM erfolgen oder zuerst ueber eine Zwischen-Datei?
6. Welche Excel-Version wird verwendet?
7. Darf VBA verwendet werden oder soll moeglichst ohne Makros gearbeitet werden?

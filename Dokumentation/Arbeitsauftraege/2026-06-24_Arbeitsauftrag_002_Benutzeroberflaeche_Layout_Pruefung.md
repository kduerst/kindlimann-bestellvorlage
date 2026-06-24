# Arbeitsauftrag 002 - Benutzeroberflaeche und Layout-Pruefung

## Repository

`kduerst/kindlimann-bestellvorlage`

## Abhaengigkeit

Dieser Arbeitsauftrag darf erst ausgefuehrt werden, wenn Arbeitsauftrag 001 erfolgreich abgeschlossen und von Kevin freigegeben wurde.

Voraussetzung:

```text
Dokumentation/Entscheidungen/Arbeitsauftrag_001_Repository_Struktur_Kindlimann_Bestellliste.md
```

muss vorhanden sein und bestaetigen, dass die Repository-Struktur sauber geordnet ist.

## Ziel

Die Benutzeroberflaeche der Kindlimann Bestellliste soll geprueft und gezielt verbessert werden, damit die Excel-Datei sauber, ruhig, modern und verlaesslich wirkt.

Der Fokus liegt auf dem Erscheinungsbild und der Bedienbarkeit der bestehenden Excel-Datei. Die vorhandenen Funktionen sollen erhalten bleiben.

## Ausgangslage

Die Kindlimann Bestellliste funktioniert grundsaetzlich. Die Oberflaeche soll jedoch optisch und praktisch verbessert werden, damit die Tabelle nicht gedrueckt oder unruhig wirkt.

Insbesondere soll die Tabelle nicht oben links an der Ecke kleben, sondern mit einem sauberen Rand platziert werden.

## Nicht-Ziel

Dieser Auftrag ist kein Funktionsumbau.

Nicht erlaubt sind:

- keine fachliche Aenderung am Bestellprozess
- keine Entfernung bestehender Funktionen
- keine Vereinfachung durch Weglassen benoetigter Funktionen
- keine unnoetige Verlaengerung oder Verkomplizierung des Codes
- keine Aenderung an Profilen, Empfaengern oder Bestelllogik, sofern nicht technisch zwingend fuer die Layoutpruefung erforderlich
- keine grossen Refactorings ohne separaten Arbeitsauftrag

## Gewuenschte Layout-Aenderungen

### Abstand um die Tabelle

Rund um die eigentliche Bestelltabelle soll ein optisch sauberer Abstand vorgesehen werden.

Richtwert:

```text
ca. 20 mm Abstand um die Tabelle
```

Ziel:

- Tabelle klebt nicht oben links in der Ecke.
- Oberflaeche wirkt geloest und hochwertiger.
- PDF-/Druckansicht bleibt sauber nutzbar.

### Spaltenbreiten und Zeilenhoehen

Alle Zeilen und Spalten sollen so dimensioniert werden, dass die Inhalte sichtbar und lesbar sind.

Grundregel:

- Spaltenbreite richtet sich am laengsten relevanten Inhalt aus.
- Kurze Werte duerfen dadurch in einer etwas breiteren Spalte stehen.
- Nichts soll gequetscht wirken.
- Nichts soll unnoetig uebergross wirken.
- Etwas zusaetzlicher Abstand ist gewuenscht, damit die Oberflaeche ruhig wirkt.

Besonders zu pruefen:

- Auftrag / Projekt
- Profil
- Laenge
- Menge
- Bemerkung
- Besteller
- Druckdatum / Bestelldatum
- Liefertermin

## Funktionspruefung

Im Rahmen der Layoutarbeit sollen bestehende Bedienfunktionen geprueft werden.

Zu pruefen:

- Pluszeichen bei den Auftraegen zum Hinzufuegen von Profilen funktionieren wie vorgesehen.
- Profil-Dropdowns funktionieren sauber.
- Besteller-Dropdown funktioniert sauber.
- Pflichtfelder und Kontrollanzeigen funktionieren wie vorgesehen.
- PDF-Export funktioniert nach der Layoutanpassung weiterhin.
- Tabellenbereiche und Druckbereiche bleiben korrekt.
- Farbgestaltung ist stimmig, schlicht und professionell.

## Code-Pruefung

Der bestehende Code soll auf Verlaesslichkeit und Verstaendlichkeit geprueft werden.

Ziel:

- Code bleibt so kurz wie sinnvoll moeglich.
- Keine kuenstliche Verlaengerung.
- Keine unnoetige Komplexitaet.
- Keine Funktionskuerzung.
- Keine Verhaltensaenderung ohne ausdrueckliche Notwendigkeit.

Wenn Auffaelligkeiten gefunden werden, sollen diese dokumentiert werden. Nur kleine, klar begrenzte Korrekturen duerfen umgesetzt werden, wenn sie direkt zur Benutzeroberflaeche oder Layoutzuverlaessigkeit gehoeren.

## Vorgehen

1. Pruefen, ob Arbeitsauftrag 001 abgeschlossen und freigegeben ist.
2. Aktuelle Excel-/VBA-/Generator-Struktur lesen.
3. Relevante Layoutstellen identifizieren.
4. Abstand um die Tabelle technisch sinnvoll umsetzen.
5. Spaltenbreiten und Zeilenhoehen anhand der Inhalte pruefen und anpassen.
6. Dropdowns und Pluszeichen-Funktion pruefen.
7. PDF-/Druckansicht pruefen.
8. Nur notwendige Layout- oder kleine Zuverlaessigkeitskorrekturen vornehmen.
9. Ergebnis dokumentieren.

## Pruefkriterien

Der Auftrag ist erfuellt, wenn:

- die Tabelle optisch nicht mehr oben links klebt
- rund um die Tabelle ein sauberer Abstand vorgesehen ist
- Spaltenbreiten und Zeilenhoehen lesbar und stimmig sind
- Dropdowns funktionieren
- Pluszeichen zum Hinzufuegen von Profilen funktionieren
- PDF-Export weiterhin funktioniert
- keine benoetigte Funktion entfernt wurde
- keine unnoetige Code-Komplexitaet hinzugefuegt wurde
- alle Aenderungen nachvollziehbar dokumentiert sind

## Erwartetes Ergebnisdokument

Nach Abschluss ist folgendes Dokument zu erstellen:

```text
Dokumentation/Entscheidungen/Arbeitsauftrag_002_Benutzeroberflaeche_Layout_Pruefung.md
```

Dieses Ergebnisdokument muss enthalten:

- Ziel
- Voraussetzung / Status Arbeitsauftrag 001
- Vorgehen
- gepruefte Bereiche
- umgesetzte Aenderungen
- bewusst nicht geaenderte Bereiche
- Testergebnis
- Blockaden
- Entscheidung
- Empfehlung fuer den naechsten Arbeitsauftrag

## Freigabe

Die finale Freigabe erfolgt durch Kevin.

Codex darf Aenderungen vorbereiten, pruefen und dokumentieren. Die Aenderungen gelten erst nach Kevins Pruefung als freigegeben.

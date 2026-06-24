# Arbeitsauftrag 001 - Repository-Struktur Kindlimann Bestellliste

## Repository

`kduerst/kindlimann-bestellvorlage`

## Ziel

Die bestehende Ablage der Kindlimann Bestellliste soll strukturell aufgeraeumt werden, damit das Repository langfristig als sicherer, nachvollziehbarer und gut wartbarer Ablagepunkt verwendet werden kann.

Der Fokus liegt ausschliesslich auf der Ordner- und Dateistruktur. Die bestehende Funktion der Excel-Datei und der VBA-/Makro-Code bleiben unveraendert.

## Ausgangslage

Die Kindlimann Bestellliste funktioniert bereits grundsaetzlich gut. Damit kuenftige Anpassungen sicherer umgesetzt werden koennen, soll das Repository so gegliedert werden, dass einzelne Bereiche klar voneinander getrennt sind.

Ziel ist, dass kleine Aenderungen gezielt an der richtigen Stelle erfolgen koennen und nicht versehentlich groessere oder fachlich nicht betroffene Teile veraendert werden.

## Nicht-Ziel

In diesem Arbeitsauftrag duerfen keine fachlichen oder technischen Codeaenderungen vorgenommen werden.

Nicht erlaubt sind insbesondere:

- keine Aenderung am VBA-/Makro-Code
- keine Aenderung an der produktiven Excel-Logik
- keine Aenderung an bestehenden Formeln
- keine Aenderung am PDF-Export-Verhalten
- keine Aenderung an Mailtexten oder Empfaengern
- keine Aenderung an hinterlegten Profilen
- keine optische Ueberarbeitung der Excel-Datei
- keine funktionale Erweiterung

Dieser Auftrag ist ein reiner Struktur- und Ordnungsauftrag.

## Gewuenschte Strukturprinzipien

Die Repository-Struktur soll sich an der etablierten Arbeitsweise aus dem bestehenden GLNDR-Projekt orientieren, jedoch nur fuer dieses Repository umgesetzt werden.

Kernprinzipien:

- Arbeiten erfolgen ueber klar benannte Arbeitsauftraege.
- Ergebnisse werden separat als Entscheidungen dokumentiert.
- Produktive Dateien, Entwicklungsdateien, Konfigurationen und Dokumentation werden getrennt.
- Dateien sollen in sinnvolle Groessen und Verantwortungsbereiche aufgeteilt werden.
- Keine unnoetig grossen Sammeldateien, wenn Inhalte logisch getrennt werden koennen.
- Keine unnoetigen Umbenennungen produktiver Dateien.
- Keine generierten Test- oder Exportdateien ins Repository aufnehmen.

## Vorgeschlagene Zielstruktur

```text
/
├─ Kindlimann Bestellliste.xlsm
├─ Bestellt/
├─ _Entwicklung/
│  ├─ Config/
│  ├─ VBA/
│  ├─ Generator/
│  └─ Tests/
├─ Dokumentation/
│  ├─ Arbeitsauftraege/
│  └─ Entscheidungen/
├─ AGENTS.md
├─ README.md
└─ .gitignore
```

## Vorgehen

1. Aktuelle Repository-Struktur aufnehmen.
2. Bestehende Dateien nach Funktion klassifizieren:
   - produktive Excel-Datei
   - Entwicklungsdateien
   - Konfiguration
   - Dokumentation
   - generierte Ausgaben
3. Zielstruktur festlegen.
4. Nur Dateien verschieben oder Dokumentationsordner ergaenzen, wenn dadurch keine Funktion veraendert wird.
5. Bestehende Datei- und Ordnernamen nur dann anpassen, wenn es fuer die Ordnung zwingend notwendig ist.
6. Nach der Strukturarbeit ein Ergebnisdokument unter `Dokumentation/Entscheidungen/` erstellen.

## Pruefkriterien

Der Arbeitsauftrag ist erfuellt, wenn:

- das Repository klarer gegliedert ist
- Arbeitsauftraege und Entscheidungen getrennt abgelegt sind
- produktive Dateien und Entwicklungsdateien unterscheidbar sind
- keine Code- oder Excel-Funktion geaendert wurde
- die bestehende Excel-Datei weiterhin unveraendert nutzbar bleibt
- die Struktur fuer spaetere kleine Aenderungen besser geeignet ist

## Erwartetes Ergebnisdokument

Nach Abschluss ist folgendes Dokument zu erstellen:

```text
Dokumentation/Entscheidungen/Arbeitsauftrag_001_Repository_Struktur_Kindlimann_Bestellliste.md
```

Dieses Ergebnisdokument muss enthalten:

- Ziel
- Vorgehen
- geaenderte Struktur
- nicht geaenderte Bereiche
- Ergebnis
- Blockaden
- Entscheidung
- Empfehlung fuer den naechsten Arbeitsauftrag

## Freigabe

Die finale Freigabe erfolgt durch Kevin.

Codex darf die Struktur vorbereiten und dokumentieren, aber keine produktive Funktionsaenderung als abgeschlossen betrachten, bevor Kevin sie geprueft hat.

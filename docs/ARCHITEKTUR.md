# Architektur und Projektstruktur

Dieses Dokument beschreibt die logische Struktur der Kindlimann-Bestellvorlage. Es dient als sicherer Orientierungspunkt fuer spaetere Arbeiten mit ChatGPT, Codex und GitHub.

## Ziel

Die Excel-Arbeitsmappe soll nicht als unuebersichtliche Einzeldatei betrachtet werden, sondern als Ergebnis aus klar getrennten Projektbestandteilen:

- produktive Excel-Datei fuer die Nutzung
- Entwicklungsdateien fuer Erzeugung und Wartung
- Konfiguration fuer Profile, Layout, Mail und Ausgabe
- Dokumentation fuer Anforderungen, Aenderungen und naechste Arbeiten
- Ausgabeordner fuer erzeugte PDFs

## Grundsatz

Keine fachliche Logik soll zufaellig geaendert werden. Aenderungen sollen gezielt an der passenden Stelle erfolgen:

- Stammdaten und Parameter in Konfigurationsdateien
- Excel-/VBA-Erzeugung im Entwicklungsbereich
- Dokumentation im Ordner `docs`
- erzeugte PDFs im Ordner `Bestellt`

## Aktuelle Hauptbereiche

```text
/
├─ Kindlimann Bestellliste.xlsm          # produktive Excel-Arbeitsmappe
├─ Bestellt/                             # lokaler Zielordner fuer erzeugte PDFs
├─ _Entwicklung/                         # technische Erstellung und Wartung
│  └─ Config/                            # zentrale Einstellungen
└─ docs/                                 # Dokumentation und Steuerung weiterer Arbeiten
```

## Empfohlene Arbeitsweise

1. Anforderungen zuerst in `docs/ANFORDERUNGEN.md` erfassen.
2. Technische Aenderung in einem separaten Branch vorbereiten.
3. Keine Misch-Aenderungen: Struktur, Layout, Makro-Logik und fachliche Regeln getrennt behandeln.
4. Nach jeder Aenderung `docs/CHANGELOG.md` pflegen.
5. Vor produktiver Nutzung die erzeugte Excel-Datei lokal in Excel Desktop testen.

## Keine Codeaenderung in diesem Schritt

Dieser Struktur-Schritt aendert keine VBA-Logik und keine bestehende Excel-Funktion. Er fuegt nur Dokumentation und klare Ablagepunkte hinzu.

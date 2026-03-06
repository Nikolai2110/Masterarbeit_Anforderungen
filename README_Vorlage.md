# Masterarbeit Word-Vorlage – Anleitung

## Voraussetzungen

Python 3.8+ muss installiert sein.

Abhängigkeiten installieren:

```bash
pip install -r requirements.txt
```

## Vorlage generieren

```bash
python generate_template.py
```

Die Datei `Masterarbeit_Vorlage.docx` wird im aktuellen Verzeichnis erstellt.

## Inhalt der generierten Datei

Die Vorlage enthält:

- **Deckblatt** mit Logo der Hochschule Osnabrück, Titel, Autor- und Prüferangaben
- **Abstract** (Deutsch und Englisch) mit Platzhaltertexten
- **Inhaltsverzeichnis** (Platzhalter – muss in Word aktualisiert werden)
- **Abkürzungsverzeichnis** mit Beispielabkürzungen
- **Abbildungsverzeichnis** und **Tabellenverzeichnis**
- **Vollständige Kapitelstruktur** (Kapitel 1–8) mit allen Unterkapiteln gemäß Exposé
- **Literaturverzeichnis**, **Anhang** (A–D) und **Eidesstattliche Erklärung**

Alle Formatvorlagen (Styles) für Überschriften, Fließtext, Fußnoten und Blockzitate sind korrekt konfiguriert.

## Hinweise nach der Generierung

### Inhaltsverzeichnis aktualisieren
Das Inhaltsverzeichnis muss in Word manuell aktualisiert werden:
1. Rechtsklick auf den Inhaltsverzeichnis-Platzhalter
2. „Felder aktualisieren" → „Gesamtes Verzeichnis aktualisieren"

### Silbentrennung aktivieren
Silbentrennung kann nicht programmatisch gesetzt werden und muss in Word aktiviert werden:
1. Reiter „Layout" → „Silbentrennung" → „Automatisch"

### Logo
Das Script versucht, das Logo der Hochschule Osnabrück automatisch herunterzuladen.
Falls der Download fehlschlägt (kein Internet, geblockte URL), wird ein Text-Fallback verwendet.
In diesem Fall kann das Logo manuell in Word ersetzt werden:
1. Den Text-Fallback auf dem Deckblatt markieren und löschen
2. Reiter „Einfügen" → „Bilder" → Logo-Datei auswählen

### Platzhalter ersetzen
Folgende Platzhalter müssen in der Datei manuell ersetzt werden:
- `[Matrikelnummer]` → Ihre Matrikelnummer
- `[Datum]` → Abgabedatum

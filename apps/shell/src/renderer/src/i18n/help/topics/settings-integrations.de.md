# Einstellungen, Sprache, Design und MCP-Integrationen

## Einstellungen öffnen

Die Kontenzeile unten links auf der Startseite öffnet den Einstellungsbereich (wenn Sie abgemeldet sind, steht dort „Anmelden“); KI-bezogene Optionen finden sich im Abschnitt „KI-Modell“

![Das Einstellungsfenster](img/settings-integrations.png) — die Modellkonfiguration ist in „KI-Modelle und Einstellungen“ behandelt.

## Sprache

- Die Einstellungen bieten **21 Oberflächensprachen**: Englisch, vereinfachtes Chinesisch, Japanisch, Koreanisch, Französisch, Deutsch, Spanisch, Thai, Indonesisch, Russisch, Arabisch, Portugiesisch, Italienisch, Polnisch, Tschechisch, Niederländisch, Malaiisch, Hebräisch, Hindi, traditionelles Chinesisch, Vietnamesisch.
- Die Umstellung gilt sofort und bleibt erhalten; die native Menüleiste wird in der Sprache neu aufgebaut.

## Design

Hell / Dunkel / System folgen. „System folgen“ richtet sich nach dem Erscheinungsbild des Betriebssystems, und die Editoren wechseln synchron die Darstellung, ohne zu flackern.

## Standard-App-Zuordnungen

Die Einstellungen können GenOffice als Programm für .docx / .xlsx / .pptx / .pdf und verwandte Formate hinterlegen (Registrierung als Standard-App auf Plattformebene; bestätigen Sie, wenn Sie dazu aufgefordert werden).

## Hinweise zu Drittanbietersoftware und Updates

- Hilfe ▸ Hinweise zu Drittanbietersoftware: das vollständige Verzeichnis der mitgelieferten Open-Source-Lizenzen.
- Hilfe ▸ Nach Updates suchen: löst eine manuelle Prüfung aus; bei einer neueren Version werden Sie zur Installation aufgefordert.

## Bei Genspark anmelden

- Der Anmeldeeinstieg (Einstellungen oder Liste der Cloud-Projekte) nutzt ein **Gerätecode**-Verfahren: GenOffice zeigt einen Code und öffnet die Anmeldung im Browser; danach geht es automatisch weiter.
- Die Anmeldung wird nur genutzt für: die Liste der Cloud-Projekte und die gehosteten Modelle von Genspark. Ohne sie funktionieren alle lokalen Funktionen und benutzerdefinierten Modelle weiter.
- Das Abmelden ist in den Einstellungen ein Klick.

## MCP-Integration (für erfahrene Anwender / KI-Clients)

GenOffice betreibt einen lokalen **MCP-Server**, damit externe KI-Clients (Claude Desktop, Cursor, …) Ihre Dokumente direkt lesen und schreiben können:

- Start: `genoffice mcp` in der Kommandozeile (Port und Authentifizierungstoken konfigurierbar; standardmäßig nur auf Loopback).
- Fähigkeiten: docx, xlsx und pptx erstellen/öffnen/bearbeiten, Inhalte lesen, Formate konvertieren, als PDF exportieren und mehr — derselbe Werkzeugsatz wie in den Desktop-Apps.
- Sicherheit: Token-Authentifizierung ist optional, aber empfohlen; der Listener bleibt standardmäßig auf dem lokalen Rechner; siehe `genoffice mcp --help`.

## Spickzettel für die Kommandozeile

| Befehl             | Was er bewirkt             |
| ------------------ | -------------------------- |
| `genoffice <file>` | eine Datei öffnen          |
| `genoffice mcp`    | lokalen MCP-Server starten |
| `genoffice --help` | alle Befehle und Optionen  |

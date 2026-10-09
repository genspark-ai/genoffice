# Einstellungen, Sprache, Design und MCP-Integrationen

## Einstellungen öffnen

Die Kontenzeile unten links auf der Startseite öffnet den Einstellungsbereich (wenn Sie abgemeldet sind, steht dort „Anmelden“). Er hat sechs Abschnitte: Konto, KI-Modell, KI-Medien & Suche, Allgemein, Integrationen und Über.

![Einstellungen ▸ Allgemein, wo Sprache, Design, AutoSave und der Schalter für die Nutzungsstatistik liegen](img/settings-general.png)

Die Modellkonfiguration hat einen eigenen Artikel; unter **KI-Medien & Suche** schalten Sie je nach Anbieter Bildgenerierung, Bildanalyse, Videoanalyse, Websuche und lokale Dateisuche ein.

## Sprache

- Die Einstellungen bieten **21 Oberflächensprachen**: Englisch, vereinfachtes Chinesisch, Japanisch, Koreanisch, Französisch, Deutsch, Spanisch, Thai, Indonesisch, Russisch, Arabisch, Portugiesisch, Italienisch, Polnisch, Tschechisch, Niederländisch, Malaiisch, Hebräisch, Hindi, traditionelles Chinesisch, Vietnamesisch.
- Die Umstellung gilt sofort und bleibt erhalten; die native Menüleiste wird in der Sprache neu aufgebaut.

## Design

Hell / Dunkel / System folgen. „System folgen“ richtet sich nach dem Erscheinungsbild des Betriebssystems, und die Editoren wechseln synchron die Darstellung, ohne zu flackern.

## Allgemein

- **Anonyme Nutzungsstatistiken senden** — standardmäßig aktiviert. Verwendet Google Analytics 4 und sendet Ihre öffentliche IP-Adresse und Transportmetadaten; Dokumentinhalte und Dateinamen werden nie erhoben, und jedes Ereignis trägt nur einen Typ wie „eine .docx geöffnet“. Sie können das hier jederzeit abschalten.
- **Position der KI-Seitenleiste** (links oder rechts), **Textgröße im KI-Bereich** und **Rechtschreibprüfung im KI-Chat**.
- **KI-Panel in neuen Dokumenten öffnen** — aus, ein neues Dokument startet mit eingeklapptem Panel, einen Klick entfernt.
- **Alle Dokumente automatisch speichern** schaltet AutoSave in jedem Editor standardmäßig ein; für ein einzelnes Fenster können Sie es weiterhin abschalten.
- **Speicherort** mit der Schaltfläche **Ändern** und **Standard-App für Office-Dokumente**, um .docx / .xlsx / .pptx für GenOffice zu beanspruchen.

## KI-Medien & Suche

Keine Schalter — jede Fähigkeit wählt den Anbieter, der sie bedient, und Schlüssel und Basis-URL eines Anbieters werden einmal eingetragen und gemeinsam genutzt:

- **Websuche**, **Bildgenerierung**, **Bildanalyse** und **Videoanalyse**, jede mit Anbieter, Modell, Schlüssel und Basis-URL.
- **Lokale Dateisuche** läuft auf diesem Rechner. Darunter ist **Jev-Neusortierung**, die **standardmäßig aus** ist. Schalten Sie sie ein, werden die Auszüge der 20 besten lokalen Treffer — bis zu 1.200 Zeichen je Dokument, dazu die Datei- und Ordnernamen — an das Jev-Modell von TypeSafe gesendet und nach Relevanz neu sortiert. Ist sie aus, verlässt nichts dieses Gerät.

## Über

- **Version**, der GitHub-Link des Projekts und eine Schaltfläche **Auf GitHub Stern geben**.
- **Update-Kanal**: Stabil oder Beta. Eine Änderung wirkt sofort und prüft auf ein Update; eine Beta-Installation wird nicht auf Stabil zurückgestuft.

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

**Integrationen** ist der Bereich, der GenOffice mit einem Coding-Agenten verbindet, und er hat einen eigenen Artikel: Einen Coding-Agenten anbinden. Die Kurzfassung — wählen Sie einen Weg (die Kommandozeile oder MCP), folgen Sie dem jeweiligen Abschnitt und starten Sie dann einen neuen Chat.

![Einstellungen ▸ Integrationen: die drei Schritte, dann die Skill-Zeilen und die MCP-Optionen](img/settings-integrations.png)

Unter **Lokaler HTTP-Server** kann die App den Server auch selbst betreiben — ein Schalter zum Aktivieren und ein Port —, während **Erweitert** die Zustandsprüfung-URL und die Protokolldatei hinzufügt, statt ihn dem Assistenten zu überlassen. Er lauscht nur auf localhost.

## Spickzettel für die Kommandozeile

| Befehl             | Was er bewirkt             |
| ------------------ | -------------------------- |
| `genoffice <file>` | eine Datei öffnen          |
| `genoffice mcp`    | lokalen MCP-Server starten |
| `genoffice --help` | alle Befehle und Optionen  |

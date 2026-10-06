# KI-Modelle und Einstellungen

## Anbieter und Modelle

Modelle und Schlüssel werden in den Einstellungen konfiguriert (die Kontenzeile unten links auf der Startseite):

![Das Einstellungsfenster](img/settings-general.png)

- **Genspark gehostet**: anmelden (Gerätecode-Verfahren) und loslegen — ganz ohne Konfiguration.
- **Eigene Endpunkte (BYOK)**: Einstellungen ▸ KI erwartet eine Basis-URL und einen API-Schlüssel pro Protokoll — OpenAI-kompatibel, Anthropic, Gemini, DeepSeek, DashScope (qwen) und mehr. Schlüssel leben nur in den Anfrage-Headern — niemals auf der Festplatte, in Protokollen oder in der Umgebung von Unterprozessen.
- Je nach Fähigkeit lässt sich ein anderes Modell wählen: Chat/Erzeugung, Bilderzeugung, Bildanalyse.
- **Verbindung testen**: prüft die Erreichbarkeit des Endpunkts und die Sichtbarkeit des Modells, bevor gespeichert wird.
- Basis-URLs dürfen Pfad und Abfragezeichenfolge enthalten (Gateway-Stil); Endpunktpfade werden korrekt angehängt.

## CLI-Integration (Codex-Klasse)

- Die Einstellungen nehmen einen Pfad zu einem lokalen CLI-Programm auf (Heimatverzeichnisse mit Nicht-ASCII-Zeichen und ein ~-Präfix funktionieren; ~ wird automatisch aufgelöst); Modelle erkennen tastet die verfügbaren Modelle der CLI ab.
- Die Prüfung kontrolliert nur die Existenz — keine Einschränkung des Zeichensatzes.

## Wann Änderungen wirksam werden

- Änderungen an Modell und Endpunkt gelten sofort; eine laufende Unterhaltung behält die alte Konfiguration bis zu ihrem nächsten Durchgang.

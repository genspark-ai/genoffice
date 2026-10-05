# Instellingen, taal, thema en MCP-integraties

## Instellingen openen

De accountregel linksonder op Start opent het instellingenpaneel (als je niet bent aangemeld staat er Inloggen); AI-gerelateerde opties vind je in de sectie AI-model

![Het venster Instellingen](img/settings-integrations.png) — modelconfiguratie wordt behandeld in AI-modellen en instellingen.

## Taal

- De instellingen bieden **21 interfacetalen**: Engels, vereenvoudigd Chinees, Japans, Koreaans, Frans, Duits, Spaans, Thais, Indonesisch, Russisch, Arabisch, Portugees, Italiaans, Pools, Tsjechisch, Nederlands, Maleis, Hebreeuws, Hindi, traditioneel Chinees, Vietnamees.
- Wisselen gaat meteen in en blijft behouden; de systeemenubalk wordt in de nieuwe taal opgebouwd.

## Thema

Licht / Donker / Systeem volgen. Systeem volgen volgt het uiterlijk van het besturingssysteem en de editors veranderen mee zonder te flikkeren.

## Standaardapp-koppelingen

Met de instellingen kun je GenOffice registreren als verwerker van .docx / .xlsx / .pptx / .pdf en vergelijkbare bestanden (registratie van de standaardapp op systeemniveau; bevestiging wanneer daarom wordt gevraagd).

## Mededelingen over software van derden en updates

- Help ▸ Kennisgevingen over software van derden: het volledige overzicht van opensourcelicenties dat met de app is meegeleverd.
- Help ▸ Controleren op updates…: start een handmatige controle; bij een nieuwere versie wordt je gevraagd die te installeren.

## Aanmelden bij Genspark

- Het aanmeldpunt (in de instellingen of in de lijst met cloudprojecten) gebruikt een flow met **apparaatcode**: GenOffice toont een code en opent de aanmeldpagina in de browser; daarna gaat het automatisch verder.
- Aanmelden wordt alleen gebruikt voor: de lijst met cloudprojecten en de modellen die Genspark host. Zonder aanmelden blijven alle lokale functies en eigen modellen werken.
- Uitloggen is in de instellingen met één klik te doen.

## MCP-integratie (voor gevorderde gebruikers / AI-clients)

GenOffice bevat een lokale **MCP-server**, zodat externe AI-clients (Claude Desktop, Cursor, ...) je documenten rechtstreeks kunnen lezen en bewerken:

- Starten: `genoffice mcp` op de opdrachtregel (poort en auth-token zijn instelbaar; standaard alleen op loopback).
- Mogelijkheden: docx-, xlsx- en pptx-bestanden maken/openen/bewerken, inhoud lezen, formaten converteren, naar pdf exporteren en meer — dezelfde gereedschapsset die de desktopapps gebruiken.
- Beveiliging: tokenauthenticatie is optioneel maar aan te raden; de listener blijft standaard op de lokale machine; zie `genoffice mcp --help`.

## Commandoregel in één oogopslag

| Opdracht           | Wat het doet                     |
| ------------------ | -------------------------------- |
| `genoffice <bestand>` | opent een bestand             |
| `genoffice mcp`    | start de lokale MCP-server       |
| `genoffice --help` | alle opdrachten en opties        |

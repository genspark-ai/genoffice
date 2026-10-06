# Instellingen, taal, thema en MCP-integraties

## Instellingen openen

De accountregel linksonder op Start opent het instellingenpaneel (als je niet bent aangemeld staat er Inloggen). Het paneel heeft zes secties: Account, AI-model, AI-media en zoeken, Algemeen, Integraties en Over.

![Instellingen ▸ Algemeen, waar taal, thema, automatisch opslaan en de schakelaar voor gebruiksstatistieken staan](img/settings-general.png)

Modelconfiguratie heeft een eigen artikel; onder **AI-media en zoeken** zet je per provider beeldgeneratie, afbeeldingsanalyse, video-analyse, zoeken op het web en lokale bestanden zoeken aan.

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

**Integraties** is het paneel dat GenOffice met een coding agent verbindt, en daarvoor is er een eigen artikel: Verbinden met een coding agent. De korte versie — kies een route (de opdrachtregel, of MCP), volg dat onderdeel en start daarna een nieuw gesprek en stel je vraag.

![Instellingen ▸ Integraties: de drie stappen, daarna de skill-rijen en de MCP-opties](img/settings-integrations.png)

Onder **Lokale HTTP-server** kan de app de server ook zelf draaien — een inschakelschakelaar en een poort —, terwijl **Geavanceerd** de URL voor de statuscontrole en het logbestand toevoegt, in plaats van hem over te laten aan de assistent. Hij luistert alleen op localhost.

## Commandoregel in één oogopslag

| Opdracht              | Wat het doet               |
| --------------------- | -------------------------- |
| `genoffice <bestand>` | opent een bestand          |
| `genoffice mcp`       | start de lokale MCP-server |
| `genoffice --help`    | alle opdrachten en opties  |

# Nastavení, jazyk, motiv a integrace MCP

## Otevření nastavení

Řádek účtu vlevo dole na domovské stránce otevře panel nastavení (při nepřihlášení ukazuje Přihlásit se); možnosti týkající se AI najdete v jeho části Model AI

![Okno Nastavení](img/settings-integrations.png) — konfigurace modelů je popsána v kapitole Modely AI a nastavení.

## Jazyk

- Nastavení nabízí **21 jazyků rozhraní**: angličtina, zjednodušená čínština, japonština, korejština, francouzština, němčina, španělština, thajština, indonéština, ruština, arabština, portugalština, italština, polština, čeština, nizozemština, malajština, hebrejština, hindština, tradiční čínština, vietnamština.
- Změna působí okamžitě a zůstává zachována; nativní pruh nabídek se znovu sestaví v novém jazyce.

## Motiv

Světlý / Tmavý / Podle systému. Volba Podle systému sleduje vzhled operačního systému a editory se přeoblekují současně bez blikání.

## Přiřazení jako výchozí aplikace

Nastavení může zaregistrovat GenOffice jako aplikaci pro soubory .docx / .xlsx / .pptx / .pdf a podobné (registrace výchozí aplikace na úrovni systému; potvrzení, když o něj systém požádá).

## Informace o softwaru třetích stran a aktualizace

- Nápověda ▸ Informace o softwaru třetích stran: úplný přehled licencí open source dodávaných s aplikací.
- Nápověda ▸ Zkontrolovat aktualizace…: spustí ruční kontrolu; novější verze vyzve k instalaci.

## Přihlášení ke Genspark

- Místo přihlášení (v nastavení nebo v seznamu cloudových projektů) používá postup s **kódem zařízení**: GenOffice zobrazí kód a otevře přihlašovací stránku v prohlížeči; po dokončení postup pokračuje automaticky.
- Přihlášení slouží pouze k: seznamu cloudových projektů a modelům hostovaným Gensparkem. Bez něj fungují dál všechny místní funkce i vlastní modely.
- Odhlášení je jedno kliknutí v nastavení.

## Integrace MCP (pro pokročilé uživatele a klienty AI)

GenOffice obsahuje místní **server MCP**, takže externí klienti AI (Claude Desktop, Cursor, ...) mohou vaše dokumenty přímo číst a upravovat:

- Spuštění: `genoffice mcp` v příkazové řádce (port a autentizační token jsou nastavitelné; ve výchozím stavu pouze smyčka).
- Možnosti: vytvoření/otevření/úprava souborů docx, xlsx a pptx, čtení obsahu, převod formátů, export do PDF a další — stejná sada nástrojů, kterou používají desktopové aplikace.
- Zabezpečení: ověření tokenem je volitelné, ale doporučené; naslouchání zůstává ve výchozím stavu na místním počítači; viz `genoffice mcp --help`.

## Přehled příkazů

| Příkaz               | Co dělá                     |
| -------------------- | --------------------------- |
| `genoffice <soubor>` | otevře soubor               |
| `genoffice mcp`      | spustí místní server MCP    |
| `genoffice --help`   | všechny příkazy a přepínače |

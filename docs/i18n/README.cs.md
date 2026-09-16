<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — open-source AI kancelářský balík: Docs, Sheets, Slides, PDF, Markdown a HTML s vestavěným AI panelem" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>První plnohodnotný open-source AI kancelářský balík na světě.</b><br>
Soubory Word, Excel, PowerPoint a PDF, upravované vámi a vaší AI, ukládané zpět ve skutečných formátech.</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="Licence: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Nejnovější verze"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Stažení"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="Hvězdy na GitHubu"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ko.md">한국어</a> · <a href="README.ja.md">日本語</a> · <a href="README.ar.md">العربية</a> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <b>Čeština</b> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>Stažení</b></a> ·
  <a href="#command-line-and-agent-skill"><b>CLI</b></a> ·
  <a href="#mcp-server"><b>MCP</b></a> ·
  <a href="https://genoffice.ai/"><b>Web</b></a> ·
  <a href="https://genoffice.ai/join"><b>Komunita</b></a> ·
  <a href="../../PRIVACY.md"><b>Soukromí</b></a>
</p>

GenOffice je bezplatná open-source alternativa k Microsoft Office pro macOS,
Windows a Linux. Otevírá a ukládá nativní soubory `.docx`, `.xlsx` a `.pptx`,
upravuje PDF, Markdown a HTML a ke každému dokumentu přidává AI agenta —
nejde o přilepené chatovací okno, ale o editor, který soubor přečte,
provede požadovanou úpravu a přesně vám ukáže, čeho se dotkl.

- **Skutečné formáty, zachování na úrovni bajtů.** Přepíše se jen to, co
  upravíte. Vše ostatní v souboru zůstává bajtově identické, takže dokumenty
  dál bezproblémově fungují ve Wordu, Excelu a PowerPointu.
- **AI, kterou můžete zkontrolovat.** Úpravy přicházejí jako sledované změny
  a rozdíly s možností vrátit je jedním kliknutím. Tabulky získávají živé
  vzorce, ne vložená čísla. Prezentace a stránky se generují přímo na plochu
  a zůstávají plně editovatelné.
- **Lokální už z principu.** Soubory se otevírají, upravují, ukládají
  i převádějí přímo na vašem počítači. PDF → Word / Excel / PowerPoint
  a Markdown → Word i HTML → Word běží celé na zařízení. Zařízení opouštějí
  jen volání AI, a to k poskytovateli, kterého si sami vyberete.
- **Vaše klíče, nebo žádné.** Přihlaste se přes Genspark a klíče nepotřebujete,
  nebo použijte vlastní klíč pro Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM,
  Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, Requesty nebo jakýkoli OpenAI
  kompatibilní endpoint, včetně lokálních serverů.
- **Skriptovatelný a připravený pro agenty.** Aplikace obsahuje příkazový
  řádek `genoffice` a skill pro Claude Code, Codex, Cursor, Gemini CLI,
  GitHub Copilot, OpenCode a Windsurf, takže kódovací agent může na vašem
  počítači vytvářet, převádět, číst a upravovat skutečné soubory Office, aniž
  by otevřel jediné okno.

**Stáhnout:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon a Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 a Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
podrobnosti a požadavky najdete v sekci [Stažení](#download).

## Ukázka

Šest aplikací, jeden AI panel a příkazový řádek pro vašeho kódovacího agenta.
Každý snímek obrazovky je ze skutečné aplikace na macOS, s AI ovládanou
promptem, který si můžete přečíst v panelu.

### 1 · Docs — otevírejte a upravujte `.docx` s AI, kterou lze zkontrolovat

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs vykresluje dvoukolonovou stránku výroční zprávy s celostránkovým titulním obrázkem, podbarvenou tabulkou KPI, záhlavím a zápatím, při 80% přiblížení se sbaleným AI panelem"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs: přehled společnosti s bannerovým obrázkem; AI zestručnila sekci Overview a vložila novou sekci s odrážkami, panel nabízí vrácení jedním kliknutím"></td>
</tr>
<tr>
<td><b>Otevře soubor přesně tak, jak jej rozvrhl Word</b> — dvoukolonové sekce, obrázky přes celou stránku, podbarvené tabulky, záhlaví a zápatí, stránkování podle Wordových metrik řádkování. Styly, komentáře, sledované změny, rovnice i ruční poznámky se zachovají beze změny.</td>
<td><b>Zadejte úpravu</b> — AI si přečte jen bloky, které potřebuje, přepíše sekci Overview a vloží novou sekci s odrážkami. Každé kolo AI je snímek, který lze vrátit; se zapnutým <b>sledováním změn</b> přicházejí úpravy jako revize ve stylu Wordu.</td>
</tr>
</table>

### 2 · Sheets — `.xlsx` s živými vzorci a grafy, ne s vloženými čísly

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets: AI přidala list Summary s tržbami podle regionu a kategorie pomocí vzorců SUMIF, plus sloupcový graf, a hlásí 43 provedených změn s tlačítkem Zpět"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets: na dotaz, který region vedl v tržbách za Q2, AI odpovídá Evropa s rozpadem podle kategorií a jako odkazy cituje buňky, které použila, vedle listu Orders"></td>
</tr>
<tr>
<td><b>Vytvořte to</b> — z jedné věty agent přidá list Summary se skutečnými vzorci <code>SUMIF</code> podle regionu a kategorie, vloží sloupcový graf a všech 43 změn provede jako jednu dávku, kterou lze vrátit.</td>
<td><b>Zeptejte se</b> — dotazy k sešitu se vracejí s odůvodněním a konkrétními buňkami použitými jako klikatelné citace. Pod kapotou: vlastní Rust engine `.xlsx`, kontingenční tabulky, filtry, podmíněné formátování a sledování vzorců.</td>
</tr>
</table>

### 3 · Slides — od promptu k prezentaci `.pptx`

<img src="../assets/readme/slides-generate.webp" alt="Časosběrný záznam GenOffice Slides generujícího investorskou prezentaci Aurora Home: AI plánuje osnovu v panelu, snímky se postupně objevují na plátně a hotová prezentace končí závěrečnou výzvou k akci" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides: titulní snímek AI vygenerované investorské prezentace Aurora Home na ploše, s původním jednořádkovým promptem a shrnutím AI o tom, co vytvořila, v panelu"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides: navržený závěrečný snímek téže 11snímkové prezentace, s pruhem miniatur vlevo a AI panelem shrnujícím dějovou linku"></td>
</tr>
<tr>
<td><b>Jeden vstupní řádek</b> — „Vytvoř 10snímkovou investorskou prezentaci pro Aurora Home…“. GenOffice naplánuje dějovou linku, vyhledá čísla a vytvoří každý snímek na ploše jako skutečný soubor `.pptx`.</td>
<td><b>Hotová prezentace na výstupu</b> — jedenáct navržených snímků s konzistentní typografií, obrázky a závěrečnou výzvou k akci; dál upravujte pomocí předloh, rozvržení, chytrých vodicích čar a nedestruktivního oříznutí, nebo požádejte panel o změnu stylu, přepsání a přeuspořádání.</td>
</tr>
</table>

### 4 · PDF — upravujte text v PDF na místě, převádějte PDF do Wordu na zařízení

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF: režim Upravit text obtahuje na stránce každý textový blok pro úpravu na místě, zatímco AI panel odpovídá na dotaz ke zprávě s citacemi stránek"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs zobrazuje dokument Word lokálně převedený z čtvrtletní zprávy Helios v PDF, otevřený na druhé záložce vedle původního PDF"></td>
</tr>
<tr>
<td><b>Upravujte přímo na stránce</b> — režim Upravit text obtahuje každý textový blok pro přepsání na místě; obsahový proud se přepisuje přes PDFium s původními fonty, ne pomocí zakrývací anotace. Zeptejte se AI na dlouhou zprávu a získejte odpovědi s citacemi stránek.</td>
<td><b>Převod na zařízení</b> — <b>PDF Converter → PDF to Word</b> vytvoří editovatelný soubor `.docx`, který se otevře v Docs vedle zdroje se zachovanými nadpisy, řádky statistik a odstavci. Cíle Excel a PowerPoint fungují stejně; skenované stránky procházejí systémovým OCR.</td>
</tr>
</table>

### 5 · HTML — AI tvůrce stránek a rozhraní, nejdřív návrhový brief

Řekněte, k čemu stránka slouží a pro koho je určená. AI nejprve navrhne
**návrhový brief** — hook, paletu, typografii a stylové směry — a teprve
podle těchto tokenů vytvoří jeden samostatný soubor `.html`.

<img src="../assets/readme/html-restyle-motion.webp" alt="Časosběrný záznam GenOffice HTML měnícího styl vstupní stránky Lumen: jeden požadavek Restyle v panelu změní temnou stránku Midnight Studio na teplou verzi Solar Daybreak, zatímco všechny sekce a veškerý obsah zůstávají na místě" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML: vygenerovaná vstupní stránka pro solární stolní lampu ve tmavém stylu Midnight Studio, zobrazená v živém náhledu s AI panelem shrnujícím právě vytvořenou stránku"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="Tatáž vstupní stránka Lumen přestylovaná AI do teplého stylu Solar Daybreak: papírové pozadí, patkové titulky a oranžová akcentní barva, se zachováním všech sekcí a veškerého textu"></td>
</tr>
<tr>
<td><b>Vygenerováno z jednoho promptu</b> — výrazný hero, karty funkcí, ceník a formulář pro waitlist pro Lumen, vytvořené ve stylu Midnight Studio. Kliknutím na jakýkoli prvek jej přestylujete, dvojklikem upravíte text nebo přejdete na zdrojový pohled v CodeMirroru.</td>
<td><b>Stejný design, nový směr</b> — jeden požadavek na <b>Restyle</b> vymění tokeny briefu a stránka se jim přizpůsobí: teplý papír, editoriální patkové písmo, sluncově oranžová akcentní barva, nic se nepřepisuje. Prezentujte na celou obrazovku nebo exportujte jako PDF či nativní editovatelný dokument Word.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML: vygenerované rozhraní osobního dashboardu pro nezávislého designéra ve stylu teplého lnu, s levým panelem, patkovým pozdravem a čtyřmi metrickými kartami"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML: vygenerovaná datová zpráva o trhu elektromobilů ve stylu novinového listu, s patkovou hlavičkou, hlavním údajem 17,3 milionu a řádkem statistik"></td>
</tr>
<tr>
<td><b>Návrhy rozhraní</b> — startovací šablona „osobní dashboard“ promění personu ve funkční rozvržení: levý panel, pozdrav, sparkline fakturovaných hodin, karty faktur a vytížení — vše skutečné HTML, které lze předat vývojáři.</td>
<td><b>Datové příběhy</b> — startovací šablona „datová zpráva“ vytvoří editoriální noviny: patková hlavička, jedno hlavní číslo, řádek statistik oddělený linkou, vložené SVG grafy a poznámka k metodice.</td>
</tr>
</table>

### 6 · Markdown — blokový editor nad čistým `.md`, s funkcí Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown: vybraný odstavec zobrazuje vyskakovací okno Ask AI se zadanou instrukcí a návrhovými čipy jako Vylepšit, Zestručnit, Rozšířit a Opravit gramatiku, plus tlačítka Odeslat hned a Přidat do fronty"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown vykresluje dokument s poznámkami k vydání obsahující tabulku, vývojový diagram Mermaid a seznam úkolů, s úvodními prompty AI panelu vlevo"></td>
</tr>
<tr>
<td><b>Zeptejte se AI na výběr</b> — vyberte jakoukoli část textu a objeví se čip <b>Ask AI</b>: zadejte instrukci nebo vyberte návrh, odešlete ji hned, nebo zařaďte více zakotvených úprav do fronty a spusťte je najednou. Stejný vstupní bod existuje v každé aplikaci.</td>
<td><b>Vykresleno, uloženo jako čistý Markdown</b> — nadpisy, seznamy, tabulky, obrázky, blok kódu a diagramy Mermaid v blokovém editoru Tiptap, zapsané zpět jako čistý soubor `.md`, s plně lokálním exportem <b>Markdown → Word</b>.</td>
</tr>
</table>

### 7 · CLI — váš kódovací agent řídí GenOffice, na vašem počítači

GenOffice obsahuje příkazový řádek `genoffice` a agentní skill. Nainstalujte
skill a Claude Code, Codex, Cursor, Gemini CLI, GitHub Copilot, OpenCode nebo
Windsurf mohou vytvářet, převádět, číst a upravovat skutečné soubory Office
přes stejné enginy jako aplikace, aniž by otevřely jediné okno.

<img src="../assets/readme/cli-deck-in-app.webp" alt="GenOffice Slides zobrazuje osmisnímkovou prezentaci o sluneční soustavě, kterou kódovací agent vytvořil přes příkazový řádek genoffice: titulní snímek na ploše, osm miniatur vlevo a otevřený AI panel" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/cli-slides-grid.webp" alt="Osm vykreslených snímků prezentace o sluneční soustavě vedle sebe: titul, časová osa průzkumu, čtyři klíčová čísla, sloupcový graf průměrů planet, kamenné světy versus obři, 99,8 % Slunce jako hlavní číslo, mřížka čtyř obrů a závěry"></td>
<td width="50%"><img src="../assets/readme/cli-integrations.webp" alt="Nastavení GenOffice, stránka Integrace: skill genoffice nainstalovaný do Claude Code, s tlačítky Instalovat u položek Codex a Cursor"></td>
</tr>
<tr>
<td><b>Jeden prompt vašemu agentovi</b> — „Vytvoř osmisnímkovou prezentaci o sluneční soustavě.“ Agent si přečte skill, napíše stylový list, osnovu a jednu specifikaci stránky na snímek, vygeneruje obě fotografie pomocí <code>genoffice image</code> a nechá <code>genoffice slides check</code> odmítnout vše, co přetéká nebo se překrývá, než <code>genoffice create</code> sestaví soubor <code>.pptx</code> a <code>slides render</code> vrátí ke kontrole jeden PNG na snímek.</td>
<td><b>Nainstalujte jednou, v Nastavení → Integrace</b> — GenOffice vypíše kódovací agenty, které na tomto počítači najde, a zapíše skill do každého, který vyberete. Nebo si skill stáhněte jako zip, případně spusťte <code>npx skills add genspark-ai/genoffice</code>. Příkazy a celý pracovní postup najdete v sekci <a href="#command-line-and-agent-skill">Příkazový řádek a agentní skill</a>.</td>
</tr>
</table>

### 8 · MCP — stejné nástroje přes Model Context Protocol

Každý příkaz `genoffice` je zároveň nástrojem MCP. Claude Code, Claude
Desktop, Cursor a jakýkoli jiný MCP klient si mohou sami spustit `genoffice
mcp` — bez instalace skillu a bez otevřeného okna — a získat 29 nástrojů
plus referenční příručky k operacím jako zdroje (resources). Druhý,
HTTP server uvnitř aplikace umožňuje agentovi sestavit dokument Word přímo
ve viditelné záložce editoru, zatímco se díváte.

<img src="../assets/readme/mcp-deck-motion.webp" alt="Zrychlený záznam toho, jak Claude Code staví osmislajdovou investorskou prezentaci o obnovitelné energii přes MCP server genoffice: hledá obrázky a fotografie, kontroluje každou kandidátní fotografii pomocí media, deck_start zapíše stylový list a osnovu, deck_page přidává vždy jednu zkontrolovanou stránku, deck_build sestaví .pptx a slides_render vrátí obrázek každého slajdu; hotová prezentace se pak otevře v GenOffice Slides" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/mcp-deck-in-app.webp" alt="GenOffice Slides zobrazuje osmislajdovou prezentaci Renewable Energy 2026, kterou Claude Code postavil přes MCP server genoffice: titulní slajd s fotografií větrné farmy na ploše a osm miniatur vlevo"></td>
<td width="50%"><img src="../assets/readme/mcp-integrations.webp" alt="Nastavení GenOffice, stránka Integrations, část MCP: jednořádkový příkaz claude mcp add pro Claude Code, blok JSON pro Cursor, Claude Desktop a další MCP klienty, a níže možnost lokálního HTTP serveru"></td>
</tr>
<tr>
<td><b>Jeden prompt, 38 volání nástrojů, žádný shell</b> — „Vytvoř osmislajdovou investorskou prezentaci o obnovitelné energii v roce 2026, se skutečnou fotografií na titulní straně a všude, kde fotografie pomůže.“ Agent stáhne obrázky a fotografie pomocí <code>search</code>, zeptá se <code>media</code>, zda je každá kandidátní fotografie skutečnou fotografií, zavolá <code>deck_start</code> se stylovým listem a osnovou, poté <code>deck_page</code> pro každý slajd; každá stránka je zkontrolována proti osnově a paletě, než je přijata, <code>deck_build</code> sestaví <code>.pptx</code>, <code>slides_audit</code> hledá přetečení, <code>slides_render</code> vrátí PNG pro každý slajd jako obrazový obsah, na který se model může podívat, a <code>deck_replace</code> opraví tři stránky, které se mu nelíbily.</td>
<td><b>Připojte se jednou, z Nastavení → Integrace</b> — zkopírujte řádek <code>claude mcp add</code> pro Claude Code, nebo blok JSON do Cursoru, Claude Desktop či jiného MCP klienta. Možnost B zapíná lokální HTTP server pro viditelný editor Word. Obojí je popsáno v sekci <a href="#mcp-server">MCP server</a>.</td>
</tr>
</table>

## Proč GenOffice

- **Open source**, Apache-2.0, vyvíjený otevřeně na GitHubu.
- **Váš k provozování.** Nativní aplikace pro macOS, Windows a Linux; soubory
  zůstávají na vašem disku a každá úprava, uložení i konverze probíhá na
  vašem počítači.
- **Skutečné soubory Office.** Nativní `.docx`, `.xlsx` a `.pptx` se
  zachováním na úrovni bajtů: nedotčené části souboru se kopírují přesně tak,
  jak byly.
- **AI, která upravuje dokument přímo.** Sledované změny v Docs, živé vzorce
  a grafy v Sheets, prezentace kreslené přímo na plochu, každý krok AI je
  snapshot, ke kterému se lze vrátit.
- **Vlastní model, vlastní klíč.** Přihlaste se přes Genspark, nebo použijte
  klíč pro Claude, OpenAI, Gemini, DeepSeek a další, včetně lokálních serverů
  a jakéhokoli OpenAI kompatibilního endpointu.
- **PDF udělané správně.** Upravujte text přímo na stránce a převádějte PDF
  do Word, Excel nebo PowerPoint na zařízení, se systémovým OCR pro skeny.
- **Také Markdown a HTML**, se stejným panelem AI a místním exportem do Word.
- **Skriptovatelný.** Příkazový řádek `genoffice`, agentní skill a MCP
  server dávají každý engine do služeb Claude Code, Claude Desktop, Codex,
  Cursor a dalších agentů, stále na zařízení.
- **Zdarma**, pro jednotlivce i týmy.

## AI backendy

**Přihlaste se přes Genspark** a nemusíte nic nastavovat: volání modelů
prochází přes proxy Genspark (rodiny Claude, GPT a Gemini) a agenti získávají
webové a obrázkové vyhledávání, generování obrázků a analýzu
obrázků/zvuku/videa.

**Nebo použijte vlastní klíč.** Nastavení → AI nabízí Claude, OpenAI, Gemini,
DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, Requesty
a OpenCode Zen/Go, plus vlastní slot pro jakýkoli OpenAI kompatibilní
endpoint (base URL + klíč), včetně lokálních serverů s modely. Vyhledávání
a média mají vlastní poskytovatele pro každou schopnost pod **AI Media
& Search**: Serper nebo Tavily pro webové vyhledávání a OpenAI, Gemini,
Doubao/Seedream, GLM, Grok, Qwen, MiniMax nebo jakýkoli OpenAI kompatibilní
endpoint pro obrázky, generování obrázků a analýzu obrázků/videa.

Celý balík nabízí světlý, tmavý a systémový vzhled. Vzhledy mění pouze to,
co je vidět na obrazovce: exporty, tisky a uložené soubory si vždy zachovávají
vlastní barvy dokumentu.

<a id="command-line-and-agent-skill"></a>

## Příkazový řádek a agentní skill

Vše, co aplikace umí se souborem, umí příkazový řádek `genoffice`
z terminálu: prohlížet, převádět, vytvářet, číst a upravovat Word, Excel,
PowerPoint, PDF, Markdown a HTML na stejných enginech, bez grafického
rozhraní. Instaluje se spolu s GenOffice, nepotřebuje vlastní runtime
a nikdy žádný dokument nikam neodesílá. Ve spojení s přibaleným **agentním
skillem** promění kódovacího agenta v dokumentového pracovníka, který
vytváří skutečné soubory Office místo přibližných napodobenin v Markdownu.

**Funguje s:** Claude Code, Codex, Cursor, Gemini CLI, GitHub Copilot,
OpenCode a Windsurf hned po instalaci, s jakýmkoli dalším agentem, který čte
skilly, a díky [MCP serveru](#mcp-server) také s Claude Desktop a jakýmkoli
MCP klientem.

### Instalace skillu

| Jak                                    | Co se stane                                                                                                                                                                                |
| -------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Nastavení → Integrace** v aplikaci   | Vypíše agenty nalezené na tomto počítači; jedním kliknutím zapíše skill do každého, který vyberete. Když nové vydání GenOffice přinese novější skill, objeví se tlačítko **Aktualizovat**. |
| **Stáhnout jako zip** na téže stránce  | Rozložení, které claude.ai, desktopové aplikace Claude a další asistenti přijímají jako nahraný skill.                                                                                     |
| `npx skills add genspark-ai/genoffice` | Nainstaluje z tohoto repozitáře do jakéhokoli agenta podporujícího skilly.                                                                                                                 |

Pak začněte nový chat a požádejte o dokument. Skill agenta naučí, kdy sáhnout
po `genoffice`, jak si soubor přečíst před úpravou a jak zkontrolovat vlastní
práci.

### Rychlý start z terminálu

```bash
genoffice --version
genoffice info report.docx --json                  # headings and blocks; or sheets, slides, pages
genoffice convert report.md --to pdf               # md/html/docx/xlsx/pptx → pdf, pdf → docx/xlsx/pptx, …
genoffice create --type docx --from notes.md --out notes.docx
genoffice create --type xlsx --from table.json --out sales.xlsx   # "=SUM(B2:B9)" cells stay live formulas
genoffice docs read report.docx --range 0-9 --json # then `docs apply --ops edits.json` edits in place
genoffice render report.docx --out shots/          # one PNG per page, to look at what you made
genoffice open sales.xlsx                          # hand the result to the editor
```

Každý příkaz vypíše jednořádkové shrnutí, nebo s `--json` jediný objekt JSON.
Úpravy jsou atomické: odmítnutá operace nechá soubor nedotčený a vrátí se
s návodnou chybovou hláškou. `genoffice help` vypíše aktuální sadu příkazů;
úplná reference je v [packages/cli/README.md](../../packages/cli/README.md).

### Co agent ve skutečnosti spouští

Prezentace o sluneční soustavě v ukázce výše vznikla z jediného promptu
v Claude Code. Agent za ní postupoval podle fázovaného pracovního postupu ze
skillu a CLI zkontrolovalo každou fázi, než začala další:

```bash
genoffice capabilities --json                        # which cloud tools GenOffice has configured
genoffice guide slides design                        # the deck workflow and layout library
genoffice image "the eight planets in a row …" --aspect 16:9 --out deck/assets/cover.jpg
genoffice slides check deck/outline.json --json      # 8 pages, no findings
genoffice slides check deck/pages/01.json --json     # builds one slide, audits overflow and overlap
…                                                    # one page file per slide, fixed until each check is clean
genoffice create --type pptx --spec deck/pages --outline deck/outline.json --out deck/solar-system.pptx --json
genoffice slides render deck/solar-system.pptx --out deck/shots --json
genoffice slides audit deck/solar-system.pptx --json    # 8 slides, no layout issues
genoffice slides replace deck/solar-system.pptx --slide 4 --spec deck/pages/05.json --json
genoffice open deck/solar-system.pptx
```

Uvnitř `genoffice` neprobíhá žádné volání modelu: agent přemýšlí, CLI staví
a kontroluje a výsledek se otevře v GenOffice nebo PowerPointu jako obyčejný
soubor `.pptx`.

<a id="mcp-server"></a>

### MCP server

Stejné příkazy jsou dostupné jako nástroje [Model Context Protocol](https://modelcontextprotocol.io)
pro asistenty, kteří neumí spustit terminál nebo kterým byste ho raději
nedávali. Vedou k tomu dvě cesty, obě ukázané s připravenými fragmenty ke
zkopírování v **Nastavení → Integrace → MCP**:

| Cesta                                | Co to je                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |
| ------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **A · `genoffice mcp`** (doporučeno) | stdio server, který si asistent spustí sám; GenOffice nemusí být otevřené. Jeden nástroj na příkaz (`info`, `convert`, `create_docx`, `create_xlsx`, `create_pptx`, `create_pdf`, `docs_read` / `docs_apply` / `docs_check`, `sheet_*`, `slides_*`, `render`, `guide`, `search`, `image`, `media`, `open`) plus fázovaný postup sestavování prezentace `deck_start` → `deck_page` → `deck_build` → `deck_replace`. Operace, specifikace a Markdown se předávají přímo v argumentech, takže funguje i klient bez souborového systému. |
| **B · Lokální HTTP server**          | Běží uvnitř aplikace GenOffice na `http://127.0.0.1:3093/mcp` (Streamable HTTP, se starším SSE). Jeho nástroje ovládají viditelnou záložku editoru Word: `create_session`, `insert_content`, `replace_blocks`, `apply_ops`, `read_document`, `save_session` — a vy sledujete, jak dokument vzniká. Ve výchozím stavu vypnutý; zapíná se ve stejném panelu nastavení.                                                                                                                                                                 |

```bash
# Claude Code
claude mcp add --transport stdio genoffice -- genoffice mcp
```

```jsonc
// Cursor, Claude Desktop nebo jakýkoli jiný MCP klient
{ "mcpServers": { "genoffice": { "command": "genoffice", "args": ["mcp"] } } }
```

`genoffice` je tady CLI dodávané uvnitř aplikace (na macOS
`/Applications/GenOffice.app/Contents/Resources/cli/genoffice`; panel
nastavení vypíše přesnou cestu pro vaši instalaci). Server nese vlastní
instrukce k pracovnímu postupu a zpřístupňuje referenční příručky k operacím
jako zdroje `genoffice://guide/*`, takže skill není potřeba; skill a MCP
server mohou koexistovat a asistent si jeden z nich vybere. Cloudové funkce
(`search`, `image`, `media`) stále jdou přes poskytovatele nastaveného v
GenOffice; všechno ostatní běží lokálně a `GENOFFICE_ALLOWED_ROOTS` omezuje
každý nástroj na složky, které uvedete.

Prezentace o obnovitelné energii z dema výše je to, jak jeden prompt v
Claude Code s připojeným pouze MCP serverem `genoffice` vypadá ze strany
protokolu:

```text
capabilities · guide(slides, spec) · guide(slides, design)
search(query) ×4                         → IEA, BNEF and IRENA figures for the slides
search(query, images) ×7 · media(url, ask) ×7
                                         → candidate photos, each one checked to be a real photograph
deck_start(dir, style, outline)          → outline checked: 8 pages to write
deck_page(dir, 0, page) … deck_page(dir, 7, page)
                                         → each page checked against the outline and the palette; one page sent again
deck_build(dir, out)                     → renewables-2026.pptx, no image failures
slides_audit(file) · slides_render(file, out)
                                         → no layout findings; 8 PNGs come back as image content
deck_replace(dir, n, page) ×3 · slides_render(file, out)
                                         → three pages fixed after looking at the renders
```

Osmatřicet volání, zhruba třináct minut, a asistent se ani jednou nedotkl
shellu: obrázky, fotografie, průvodci, kontroly i rendery cestovaly jako
výsledky volání MCP nástrojů. Pouze `search` a `media` opustily počítač, k
poskytovateli nastavenému v GenOffice.

<a id="download"></a>

## Stažení

| Platforma                            | Požadavky                                             | Stažení                                                                                        |
| ------------------------------------ | ----------------------------------------------------- | ---------------------------------------------------------------------------------------------- |
| **macOS** — Apple Silicon (arm64)    | macOS 11+                                             | [Nejnovější `.dmg` (arm64)](https://github.com/genspark-ai/genoffice/releases/latest)          |
| **macOS** — Intel (x64)              | macOS 11+                                             | [Nejnovější `.dmg` (x64)](https://github.com/genspark-ai/genoffice/releases/latest)            |
| **Windows** (x64, většina PC)        | Windows 10+, Intel/AMD                                | [Nejnovější instalátor `-x64.exe`](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** na Arm (ARM64)           | Windows 11 na Arm (Snapdragon X a podobné)            | [Nejnovější instalátor `-arm64.exe`](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 nebo novější)       | [Nejnovější `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)                  |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [Nejnovější `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)                  |
| **Linux** — ostatní distribuce       | x86_64, glibc 2.34+, FUSE 2                           | [Nejnovější `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)             |

Všechny sestavení pocházejí z `main`; instalátory pro macOS a Windows jsou
podepsané. Starší verze jsou na stránce [Releases](https://github.com/genspark-ai/genoffice/releases).

<details>
<summary><b>Instalace na Linuxu</b></summary>

Balíček deb se instaluje pomocí apt — stáhne závislosti a přidá GenOffice
do nabídky aplikací:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

Na Fedoře / v rodině RHEL / na openSUSE nainstalujte místo toho rpm balíček:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

AppImage běží na místě: nainstalujte runtime FUSE 2
(`sudo apt install libfuse2`; na Ubuntu 24.04 se balíček jmenuje `libfuse2t64`),
udělejte soubor spustitelným a spusťte jej:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## Jak to funguje

Sedm aplikací Electron — Docs, Sheets, Slides, PDF, Markdown, HTML
a záložková shell aplikace — sdílí jednu vrstvu enginu z čistých TypeScript
balíčků plus Rust sidecar pro `.xlsx`. Původní soubor je vždy zdrojem pravdy:
úpravy se aplikují jako úzké patche a vše, čeho se editor nedotkl, přežije
zpětné uložení beze změny.

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

Prohlídka balíčků po balíčcích (enginy docx/pptx, `pdf2docx`, `html2docx`,
jádro agenta a poskytovatelé) je v [CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages).

## Vývoj

```bash
npm install
npm run fixtures     # generate test .docx fixtures
npm test             # engine + app unit tests (docs/sheets/slides need no display)
npm run typecheck    # tsc --noEmit across every workspace
npm run dev          # all six editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
npm run dist:linux   # package Linux AppImage + deb + rpm
```

Aplikace sheets navíc potřebuje pro svůj xlsx sidecar Rust toolchain
(`cargo` na PATH); `npm run build -w @genoffice/sheets` jej zkompiluje
automaticky. Kontroly, které musí projít každá změna, a postup přijímání
pull requestů najdete v [CONTRIBUTING.md](../../CONTRIBUTING.md).

## Komunita

GenOffice je v aktivním vývoji a vaše zpětná vazba jej utváří.

- **Nahlaste chybu nebo navrhněte funkci** v
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues).
- **Připojte se ke skupinovému chatu GenOffice** na
  [GenTeam](https://genoffice.ai/join), kde si můžete promluvit s týmem
  i ostatními uživateli.
- **Ohvězdičkujte repozitář**, pokud je pro vás GenOffice užitečný — je to
  nejlepší způsob, jak projekt podpořit.

## Často kladené otázky

<details>
<summary><b>Je GenOffice zdarma?</b></summary>

Ano. GenOffice je bezplatný a open-source pod licencí Apache-2.0 — žádná
zkušební verze, žádný placený tarif pro samotné aplikace.

</details>

<details>
<summary><b>Umí GenOffice otevírat soubory Microsoft Word, Excel a PowerPoint?</b></summary>

Ano. GenOffice otevírá a ukládá nativní soubory `.docx`, `.xlsx` a `.pptx`.
Ukládání zachovává bajty: části souboru, kterých jste se nedotkli, se zapíší
zpět beze změny, takže dokumenty dál fungují v Microsoft Office.

</details>

<details>
<summary><b>Funguje GenOffice offline?</b></summary>

Úprava dokumentů je zcela lokální — soubory neopustí váš počítač při
otevírání, úpravách, ukládání ani konverzi. AI funkce (agenti, vyhledávání,
nástroje pro obrázky) potřebují síťové připojení, buď s přihlášením přes
Genspark, nebo s vlastním API klíčem k modelu.

</details>

<details>
<summary><b>Umí GenOffice upravovat soubory PDF?</b></summary>

Ano — skutečná úprava textu a obrázků v PDF, která přepisuje obsahový proud
stránky se zachováním původních fontů, ne pomocí zakrývacích anotací.

</details>

<details>
<summary><b>Umí GenOffice převádět PDF do Wordu, Excelu nebo PowerPointu?</b></summary>

Ano — zcela na zařízení: extrakce na úrovni znaků přes PDFium plus analýza
rozvržení na základě geometrie, žádná cloudová služba, žádné nahrávání.
Skenované stránky jsou také podporovány: na macOS a Windows je čte systémové
OCR, takže se převedou na editovatelný text, ne na obrázek stránky.

</details>

<details>
<summary><b>Můžu použít vlastní AI model nebo API klíč?</b></summary>

Ano. Kromě bezklíčového přihlášení Genspark GenOffice podporuje vlastní klíč
pro Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok,
Mistral, OpenRouter, Requesty a OpenCode Zen/Go, plus jakýkoli OpenAI kompatibilní
endpoint — včetně lokálních serverů s modely. Vyhledávání, generování obrázků
a analýza obrázků/videa mají vlastní klíče pod Nastavení → AI Media & Search.

</details>

<details>
<summary><b>Umí GenOffice převádět HTML do Wordu?</b></summary>

Ano — Export as Word v aplikaci HTML vytvoří nativní, editovatelný soubor
`.docx` zcela na zařízení. Stránka se vykreslí ve vestavěném Chromiu
a zredukuje na skutečné struktury Wordu: nadpisy, odstavce, seznamy, tabulky,
karty, řádky KPI, pole formulářů a pozadí stránky; pouze vizuály bez
odpovídajícího prvku ve Wordu (grafy, ikony, ozdobné rámečky) se vloží jako
obrázky.

</details>

<details>
<summary><b>Můžu GenOffice ovládat z nástrojů Claude Code, Codex, Cursor nebo ze skriptu?</b></summary>

Ano. GenOffice nainstaluje příkazový řádek `genoffice`, který spouští stejné
enginy bez grafického rozhraní: prohlížejte, převádějte, vytvářejte, čtěte
a upravujte dokumenty z terminálu nebo ze skriptu, s výstupem `--json` pro
programy. Přibalený agentní skill naučí Claude Code, Codex, Cursor, Gemini
CLI, GitHub Copilot, OpenCode a Windsurf, jak ho používat; nainstalujte ho
v **Nastavení → Integrace**. Viz
[Příkazový řádek a agentní skill](#command-line-and-agent-skill).

</details>

<details>
<summary><b>Sbírá GenOffice nějaká data?</b></summary>

Oficiální balíčkované verze standardně odesílají omezenou analytiku
používání a hlášení lze kdykoli vypnout v Nastavení → Obecné. Analytika
nikdy neodesílá obsah dokumentu, názvy souborů, cesty k souborům, identitu
účtu ani e-mailové adresy. Úplný výčet událostí a zveřejnění dat najdete
v dokumentu [GenOffice Privacy](../../PRIVACY.md).

</details>

## Zabezpečení

Postoj k zabezpečení procesů (sandboxing rendereru, validace IPC, filtrování
externích odkazů) a modely hrozeb pro obsah generovaný AI najdete v
[SECURITY.md](../../SECURITY.md).

## Poděkování

GenOffice by nebyl možný bez těchto open-source projektů:

- [Electron](https://www.electronjs.org/) — desktopový runtime pro každou aplikaci.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — jádro
  uživatelského rozhraní tabulek, které Sheets rozšiřuje.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause, součástí
  balíčku přes [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer)) —
  engine obsahového proudu za skutečnou úpravou textu a obrázků v PDF.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) a
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — vykreslování PDF
  a sestavování dokumentů.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  blokové editory v Docs a Markdown.
- [CodeMirror](https://codemirror.net/) (MIT) — zdrojový editor v HTML.
- [Konva](https://konvajs.org/) — vykreslování plátna pro Slides a grafy v Sheets.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — metriky
  text-shapingu pro složitá písma.
- [calamine](https://github.com/tafia/calamine) a
  [IronCalc](https://github.com/ironcalc/IronCalc) — vrstvy pro čtení a výpočty
  v Rust xlsx sidecaru.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) — dekodér
  MicroType Express pro vložená písma PowerPointu, přenesený do
  TypeScriptu.
- [React](https://react.dev/) (MIT) — vrstva uživatelského rozhraní každé
  aplikace.
- [Mermaid](https://mermaid.js.org/) (MIT) a [KaTeX](https://katex.org/)
  (MIT) — diagramy a matematika v Markdownu a Docs.
- [opentype.js](https://opentype.js.org/) (MIT) — parsování písem pro
  metriky a vyhledávání glyfů.
- [JSZip](https://stuk.github.io/jszip/) (MIT) a
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) — vrstvy kontejneru OOXML a XML.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) — sada ikon použitá na pásech nástrojů.
- [electron-updater](https://www.electron.build/) (MIT) — aktualizace v
  rámci aplikace.
- Fonty Liberation, Carlito, Caladea a Noto CJK (OFL/Apache-2.0) — dodávané
  fonty dokumentů.

`npm run notices` znovu vygeneruje souhrn licencí dodávaných třetích stran
(`tools/gen-third-party-notices.mjs`); všechny runtime závislosti jsou
MIT/Apache-2.0/BSD-3-Clause/OFL.

## Licence

GenOffice je licencován pod [Apache License 2.0](../../LICENSE), s jednou
výjimkou: adresář `ee/` je vyhrazen pro budoucí podnikové moduly a je kryt
[GenOffice Enterprise License](../../ee/LICENSE).

Názvy a loga GenOffice a Genspark jsou ochranné známky společnosti Mainfunc,
Inc. Licence Apache-2.0 neuděluje oprávnění je používat (viz oddíl 6); forky
by měly používat vlastní branding.

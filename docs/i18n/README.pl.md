<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — otwarty pakiet biurowy AI: Docs, Sheets, Slides, PDF, Markdown i HTML z wbudowanym panelem AI" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>Pierwszy na świecie w pełni funkcjonalny otwarty pakiet biurowy z AI.</b><br>
Pliki Word, Excel, PowerPoint i PDF, edytowane przez Ciebie i Twoją AI, zapisywane z powrotem w prawdziwych formatach.</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="Licencja: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Najnowsze wydanie"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Liczba pobrań"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="Gwiazdki na GitHubie"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ko.md">한국어</a> · <a href="README.ja.md">日本語</a> · <a href="README.ar.md">العربية</a> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <b>Polski</b> · <a href="README.cs.md">Čeština</a> · <a href="README.id.md">Bahasa Indonesia</a> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>Pobierz</b></a> ·
  <a href="https://genoffice.ai/"><b>Strona internetowa</b></a> ·
  <a href="https://genoffice.ai/join"><b>Społeczność</b></a> ·
  <a href="../../PRIVACY.md"><b>Prywatność</b></a>
</p>

GenOffice to darmowa, otwarta alternatywa dla Microsoft Office na macOS,
Windows i Linux. Otwiera i zapisuje natywne pliki `.docx`, `.xlsx` i `.pptx`,
edytuje PDF, Markdown i HTML, a przy każdym dokumencie stawia agenta AI — nie
jest to czat dopięty z boku, lecz edytor, który czyta plik, wprowadza zmianę
i pokazuje precyzyjnie, czego dotknął.

- **Prawdziwe formaty, zachowane bit po bicie.** Przepisywane jest tylko to,
  co edytujesz. Cała pozostała treść pliku pozostaje nienaruszona bit po
  bicie, dzięki czemu dokumenty wciąż działają w Word, Excel i PowerPoint.
- **AI, którą możesz zweryfikować.** Zmiany trafiają jako śledzone zmiany
  i różnice z jednokrotnym przywróceniem. Arkusze kalkulacyjne dostają żywe
  formuły, a nie wklejone liczby. Prezentacje i strony są generowane na
  płótnie i pozostają w pełni edytowalne.
- **Lokalność jako zasada projektowa.** Pliki są otwierane, edytowane,
  zapisywane i konwertowane na Twoim urządzeniu. Konwersje PDF → Word / Excel
  / PowerPoint, Markdown → Word i HTML → Word działają lokalnie. Jedynie
  zapytania do AI opuszczają maszynę — i trafiają do dostawcy, którego sam
  wybierzesz.
- **Twoje klucze albo żadne.** Zaloguj się przez Genspark i pomiń klucze, albo
  użyj własnego klucza dla Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen,
  Doubao, MiniMax, Grok, Mistral, OpenRouter, API Route, Requesty lub dowolnego endpointu
  zgodnego z OpenAI, w tym lokalnych serwerów.

**Pobierz:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon i Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 i Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
szczegóły i wymagania w sekcji [Pobierz](#download).

## Demo

Sześć aplikacji, jeden panel AI. Każdy zrzut ekranu pochodzi z prawdziwej
aplikacji na macOS, a AI jest prowadzona promptem widocznym w panelu.

### 1 · Docs — otwieraj i edytuj pliki `.docx` z AI, którą możesz zweryfikować

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs renderuje dwukolumnową stronę raportu rocznego z pełnowymiarowym zdjęciem na okładce, zacienioną tabelą KPI, nagłówkiem i stopką, przy powiększeniu 80% ze zwiniętym panelem AI"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs: przegląd firmy z obrazem w formie banera; AI skróciła sekcję Overview i wstawiła nową sekcję z punktami, a panel proponuje jednokrotne przywrócenie"></td>
</tr>
<tr>
<td><b>Otwiera plik tak, jak układa go Word</b> — dwukolumnowe sekcje, obrazy na całą szerokość, zacienione tabele, nagłówki i stopki, paginacja według metryk linii Worda. Style, komentarze, śledzone zmiany, równania i pismo odręczne wracają bez zmian.</td>
<td><b>Poproś o zmianę</b> — AI czyta potrzebne fragmenty, przepisuje sekcję Overview i wstawia nową sekcję z punktami. Każdy krok AI to zrzut, który można przywrócić; przy włączonym <b>śledzeniu zmian</b> edycje trafiają jako rewizje w stylu Worda.</td>
</tr>
</table>

### 2 · Sheets — `.xlsx` z żywymi formułami i wykresami, a nie wklejonymi liczbami

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets: AI dodała arkusz Summary z przychodami według regionu i kategorii przy użyciu formuł SUMIF, plus wykres kolumnowy, i zgłasza 43 zastosowane zmiany z przyciskiem Undo"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets: na pytanie, który region przewodził w przychodach za Q2, AI odpowiada Europa z podziałem na kategorie i podaje jako klikalne odnośniki komórki, z których skorzystała, obok arkusza Orders"></td>
</tr>
<tr>
<td><b>Zbuduj to</b> — z jednego zdania agent dodaje arkusz Summary z prawdziwymi formułami <code>SUMIF</code> według regionu i kategorii, wstawia wykres kolumnowy i stosuje 43 zmiany jako jedną grupę do wycofania.</td>
<td><b>Zapytaj</b> — pytania o skoroszyt wracają z uzasadnieniem i konkretnymi komórkami użytymi jako klikalne cytowania. Pod maską: własny silnik `.xlsx` w Rust, tabele przestawne, filtry, formatowanie warunkowe i śledzenie formuł.</td>
</tr>
</table>

### 3 · Slides — od promptu do prezentacji `.pptx`

<img src="../assets/readme/slides-generate.webp" alt="Timelapse GenOffice Slides generującego prezentację inwestorską Aurora Home: AI planuje fabułę w panelu, slajdy pojawiają się na płótnie jeden po drugim, a gotowa prezentacja kończy się finałowym wezwaniem do działania" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides: slajd tytułowy wygenerowanej przez AI prezentacji inwestorskiej Aurora Home na płótnie, z oryginalnym jednowierszowym promptem i podsumowaniem AI tego, co zbudowała, w panelu"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides: zaprojektowany slajd końcowy tej samej 11-slajdowej prezentacji, z paskiem miniatur po lewej i panelem AI podsumowującym fabułę"></td>
</tr>
<tr>
<td><b>Jedna linijka na wejściu</b> — „Stwórz 10-slajdową prezentację inwestorską dla Aurora Home…”. GenOffice planuje fabułę, sprawdza dane liczbowe i tworzy każdy slajd na płótnie jako prawdziwy plik `.pptx`.</td>
<td><b>Gotowa prezentacja na wyjściu</b> — jedenaście zaprojektowanych slajdów o spójnej typografii, grafice i zamykającym wezwaniu do działania; edytuj dalej za pomocą wzorców, układów, inteligentnych linii pomocniczych i nieinwazyjnego kadrowania, albo poproś panel o zmianę stylu, przepisanie i przełożenie kolejności.</td>
</tr>
</table>

### 4 · PDF — edytuj tekst PDF w miejscu, konwertuj PDF do Worda lokalnie

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF: tryb Edit text zaznacza każdy blok tekstu na stronie do edycji w miejscu, a panel AI odpowiada na pytanie o raport z cytowaniem stron"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs wyświetla dokument Word przekonwertowany lokalnie z kwartalnego raportu Helios w formacie PDF, otwarty w drugiej karcie obok oryginalnego PDF"></td>
</tr>
<tr>
<td><b>Edytuj wewnątrz strony</b> — tryb Edit text zaznacza każdy blok tekstu do ponownego wpisania w miejscu; strumień treści jest przepisywany przez PDFium z zachowaniem oryginalnych czcionek, a nie zasłaniany anotacją. Zadaj AI pytanie o długi raport i otrzymaj odpowiedzi z cytowaniem stron.</td>
<td><b>Konwertuj lokalnie</b> — <b>PDF Converter → PDF to Word</b> tworzy edytowalny plik `.docx`, który otwiera się w Docs obok źródła, z zachowanymi nagłówkami, wierszami statystyk i akapitami. Konwersje do Excel i PowerPoint działają analogicznie; zeskanowane strony przechodzą przez systemowe OCR.</td>
</tr>
</table>

### 5 · HTML — konstruktor stron i interfejsów AI, zaczynający od briefu projektowego

Powiedz, do czego służy strona i dla kogo jest przeznaczona. AI najpierw
proponuje **brief projektowy** — hasło, paletę kolorów, typografię i kierunki
stylistyczne — a potem buduje jeden samodzielny plik `.html` zgodny z tymi
wytycznymi.

<img src="../assets/readme/html-restyle-motion.webp" alt="Timelapse GenOffice HTML zmieniającego styl strony docelowej Lumen: jedno żądanie Restyle w panelu zmienia ciemną stronę Midnight Studio w ciepłą wersję Solar Daybreak, podczas gdy każda sekcja i cała treść pozostają na miejscu" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML: wygenerowana strona docelowa dla solarnej lampki biurkowej w ciemnym kierunku Midnight Studio, wyświetlana w podglądzie na żywo z panelem AI podsumowującym właśnie zbudowaną stronę"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="Ta sama strona docelowa Lumen przestylizowana przez AI na ciepły kierunek Solar Daybreak: papierowe tło, szeryfowe nagłówki i pomarańczowy akcent, z zachowaniem wszystkich sekcji i całej treści"></td>
</tr>
<tr>
<td><b>Wygenerowana z jednego promptu</b> — odważny hero, karty funkcji, cennik i formularz listy oczekujących dla Lumen, zbudowane w kierunku Midnight Studio. Kliknij dowolny element, aby zmienić jego styl, kliknij dwukrotnie, aby edytować tekst, albo przełącz się na widok źródła w CodeMirror.</td>
<td><b>Ten sam projekt, nowy kierunek</b> — jedno żądanie <b>Restyle</b> zamienia wytyczne briefu i strona się dostosowuje: ciepły papier, edytorski szeryf, słoneczno-pomarańczowy akcent, bez przepisywania niczego od nowa. Prezentuj na pełnym ekranie albo eksportuj jako PDF lub natywny, edytowalny dokument Word.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML: wygenerowany interfejs osobistego panelu dla projektanta freelancera w ciepłym, lnianym stylu, z lewym panelem bocznym, szeryfowym powitaniem i czterema kartami metryk"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML: wygenerowany raport danych o rynku pojazdów elektrycznych w stylu gazetowym, z szeryfową winietą tytułową, nagłówkową liczbą 17,3 miliona i wierszem statystyk"></td>
</tr>
<tr>
<td><b>Makiety interfejsów</b> — starter „personal dashboard” zamienia persona w działający układ: lewy panel boczny, powitanie, wykres iskrowy godzin rozliczeniowych, karty faktur i wykorzystania — wszystko jako prawdziwy HTML, który można przekazać programiście.</td>
<td><b>Historie danych</b> — starter „data report” tworzy edytorski raport gazetowy: szeryfową winietę tytułową, jedną liczbę nagłówkową, wiersz statystyk oddzielony linią, wbudowane wykresy SVG i notę metodologiczną.</td>
</tr>
</table>

### 6 · Markdown — edytor blokowy nad zwykłym plikiem `.md`, z funkcją Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown: zaznaczony akapit pokazuje wyskakujące okno Ask AI z wpisaną instrukcją i chipami z podpowiedziami, takimi jak Polish, Make more concise, Expand i Fix grammar, a także przyciskami Send now i Add to queue"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown renderuje dokument z notatkami o wydaniu z tabelą, diagramem przepływu Mermaid i listą zadań, z podpowiedziami startowymi panelu AI po lewej stronie"></td>
</tr>
<tr>
<td><b>Zapytaj AI o zaznaczenie</b> — zaznacz dowolny fragment i pojawi się chip <b>Ask AI</b>: wpisz instrukcję lub wybierz podpowiedź, wyślij ją natychmiast albo umieść w kolejce kilka przypisanych do miejsc edycji i wykonaj je za jednym razem. Ta sama funkcja jest dostępna w każdej aplikacji.</td>
<td><b>Renderowane, zapisywane jako zwykły Markdown</b> — nagłówki, listy, tabele, obrazy, blok kodu i diagramy Mermaid w blokowym edytorze Tiptap, zapisywane z powrotem jako zwykły plik `.md`, z w pełni lokalnym eksportem <b>Markdown → Word</b>.</td>
</tr>
</table>

## Dlaczego GenOffice

- **Otwarty kod źródłowy**, Apache-2.0, tworzony w pełni otwarcie na GitHubie.
- **Twój do uruchomienia.** Natywne aplikacje na macOS, Windows i Linux; pliki
  zostają na Twoim dysku, a każda edycja, zapis i konwersja odbywają się na
  Twoim urządzeniu.
- **Prawdziwe pliki Office.** Natywne `.docx`, `.xlsx` i `.pptx`, z
  zachowaniem bit po bicie: nietknięte części pliku są kopiowane w
  niezmienionej formie.
- **AI, która edytuje sam dokument.** Śledzone zmiany w Docs, żywe formuły i
  wykresy w Sheets, prezentacje rysowane bezpośrednio na płótnie, każda tura
  AI to zapis, do którego można wrócić.
- **Twój model, Twój klucz.** Zaloguj się przez Genspark albo użyj własnego
  klucza do Claude, OpenAI, Gemini, DeepSeek i innych, z lokalnymi serwerami i
  każdym endpointem zgodnym z OpenAI włącznie.
- **PDF zrobiony właściwie.** Edytuj tekst wewnątrz strony i konwertuj PDF do
  Word, Excel lub PowerPoint lokalnie, z systemowym OCR dla skanów.
- **Także Markdown i HTML**, z tym samym panelem AI i lokalnym eksportem do
  Word.
- **Bezpłatne**, dla pojedynczych osób i zespołów.

## Silniki AI

**Zaloguj się przez Genspark** i nie musisz niczego konfigurować: wywołania
modeli trafiają przez proxy Genspark (rodziny Claude, GPT i Gemini), a
agenci dostają wyszukiwanie w internecie i obrazów, generowanie obrazów oraz
analizę obrazu, audio i wideo.

**Albo użyj własnego klucza.** Ustawienia → AI zawierają Claude, OpenAI,
Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral,
OpenRouter, API Route, Requesty oraz OpenCode Zen/Go, a także dowolny slot dla endpointu zgodnego
z OpenAI (adres bazowy + klucz), w tym lokalnych serwerów modeli.
Wyszukiwanie i media mają własnych dostawców przypisanych do konkretnej
funkcji w sekcji **AI Media & Search**: Serper lub Tavily do wyszukiwania w
internecie oraz OpenAI, Gemini, Doubao/Seedream, GLM, Grok, Qwen, MiniMax lub
dowolny endpoint obrazów zgodny z OpenAI do generowania obrazów oraz analizy
obrazu i wideo.

Cały pakiet obsługuje motywy jasny, ciemny i systemowy. Motywy zmieniają
tylko to, co widać na ekranie: eksporty, wydruki i zapisane pliki zawsze
zachowują oryginalne kolory dokumentu.

<a id="download"></a>

## Pobierz

| Platforma                               | Wymagania                                             | Pobierz                                                                                       |
| --------------------------------------- | ----------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| **macOS** — Apple Silicon (arm64)       | macOS 11+                                             | [Najnowszy plik `.dmg` (arm64)](https://github.com/genspark-ai/genoffice/releases/latest)     |
| **macOS** — Intel (x64)                 | macOS 11+                                             | [Najnowszy plik `.dmg` (x64)](https://github.com/genspark-ai/genoffice/releases/latest)       |
| **Windows** (x64, większość komputerów) | Windows 10+, Intel/AMD                                | [Najnowszy instalator `-x64.exe`](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** na Arm (ARM64)              | Windows 11 na Arm (Snapdragon X i podobne)            | [Najnowszy instalator `-arm64.exe`](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu             | x86_64, glibc 2.34+ (Ubuntu 22.04 lub nowszy)         | [Najnowszy plik `.deb`](https://github.com/genspark-ai/genoffice/releases/latest)             |
| **Linux** — Fedora / RHEL / openSUSE    | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [Najnowszy plik `.rpm`](https://github.com/genspark-ai/genoffice/releases/latest)             |
| **Linux** — inne dystrybucje            | x86_64, glibc 2.34+, FUSE 2                           | [Najnowszy plik `.AppImage`](https://github.com/genspark-ai/genoffice/releases/latest)        |

Wszystkie kompilacje pochodzą z `main`; instalatory na macOS i Windows są
podpisane. Starsze wersje znajdują się na stronie
[Releases](https://github.com/genspark-ai/genoffice/releases).

<details>
<summary><b>Instalacja na Linuksie</b></summary>

Plik deb instaluje się przez apt — ściąga zależności i dodaje GenOffice do
menu aplikacji:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

Na Fedorze / rodzinie RHEL / openSUSE zainstaluj pakiet rpm:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / rodzina RHEL
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

AppImage działa bez instalacji: zainstaluj środowisko FUSE 2
(`sudo apt install libfuse2`; na Ubuntu 24.04 pakiet nazywa się
`libfuse2t64`), nadaj plikowi prawo wykonywania, a potem go uruchom:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## Jak to działa

Siedem aplikacji Electron — Docs, Sheets, Slides, PDF, Markdown, HTML i
powłoka z kartami — współdzieli jedną warstwę silnika złożoną z czystych
pakietów TypeScript oraz sidecara w Rust dla `.xlsx`. Oryginalny plik jest
zawsze jedynym źródłem prawdy: zmiany są nakładane jako wąskie łatki, a
wszystko, czego edytor nie dotknął, wraca bez zmian.

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

Przegląd poszczególnych pakietów (silniki docx/pptx, `pdf2docx`, `html2docx`,
jądro agenta i dostawcy) znajduje się w
[CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages).

## Rozwój

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

Aplikacja sheets dodatkowo wymaga zestawu narzędzi Rust dla swojego
sidecara xlsx (`cargo` w PATH); `npm run build -w @genoffice/sheets`
kompiluje go automatycznie. Zobacz [CONTRIBUTING.md](../../CONTRIBUTING.md),
aby poznać kontrole, które musi przejść każda zmiana, oraz sposób
scalania pull requestów.

## Społeczność

GenOffice jest w aktywnym rozwoju, a Twoja opinia go kształtuje.

- **Zgłoś błąd lub poproś o funkcję** w
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues).
- **Przyłącz się do czatu grupowego GenOffice** na
  [GenTeam](https://genoffice.ai/join), aby rozmawiać z zespołem i innymi
  użytkownikami.
- **Oznacz repozytorium gwiazdką**, jeśli GenOffice jest dla Ciebie
  przydatne — to najlepszy sposób wsparcia projektu.

## FAQ

<details>
<summary><b>Czy GenOffice jest darmowe?</b></summary>

Tak. GenOffice jest darmowe i otwarte na licencji Apache-2.0 — bez okresu
próbnego, bez płatnego poziomu dla samych aplikacji.

</details>

<details>
<summary><b>Czy GenOffice może otwierać pliki Microsoft Word, Excel i PowerPoint?</b></summary>

Tak. GenOffice otwiera i zapisuje natywne pliki `.docx`, `.xlsx` i `.pptx`.
Zapis zachowuje treść bit po bicie: części pliku, których nie dotknięto, są
zapisywane z powrotem bez zmian, dzięki czemu dokumenty wciąż działają w
Microsoft Office.

</details>

<details>
<summary><b>Czy GenOffice działa offline?</b></summary>

Edycja dokumentów jest w pełni lokalna — pliki nigdy nie opuszczają Twojego
urządzenia, aby zostać otwarte, edytowane, zapisane lub przekonwertowane.
Funkcje AI (agenci, wyszukiwanie, narzędzia obrazu) wymagają połączenia
sieciowego, wraz z logowaniem przez Genspark albo własnym kluczem API
modelu.

</details>

<details>
<summary><b>Czy GenOffice może edytować pliki PDF?</b></summary>

Tak — to prawdziwa edycja tekstu i obrazów w PDF, która przepisuje strumień
treści strony z zachowaniem oryginalnych czcionek, a nie zasłania go
anotacjami.

</details>

<details>
<summary><b>Czy GenOffice może konwertować PDF do Word, Excel lub PowerPoint?</b></summary>

Tak — całkowicie lokalnie: ekstrakcja znaków na poziomie PDFium wraz z
analizą układu opartą na geometrii, bez usługi w chmurze, bez wysyłania
plików. Zeskanowane strony są również obsługiwane: na macOS i Windows
odczytuje je systemowe OCR, dzięki czemu są konwertowane do edytowalnego
tekstu, a nie obrazu strony.

</details>

<details>
<summary><b>Czy mogę użyć własnego modelu AI lub klucza API?</b></summary>

Tak. Oprócz bezkluczowego logowania przez Genspark, GenOffice wspiera
użycie własnego klucza dla Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM,
Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, API Route, Requesty i OpenCode Zen/Go, a także
dowolnego endpointu zgodnego z OpenAI — w tym lokalnych serwerów modeli.
Wyszukiwanie, generowanie obrazów oraz analiza obrazu i wideo mają własne
klucze w Ustawienia → AI Media & Search.

</details>

<details>
<summary><b>Czy GenOffice może konwertować HTML do Worda?</b></summary>

Tak — opcja Export as Word w aplikacji HTML tworzy natywny, edytowalny plik
`.docx` całkowicie lokalnie. Strona jest renderowana we wbudowanym Chromium
i redukowana do prawdziwych struktur Worda: nagłówków, akapitów, list,
tabel, kart, wierszy KPI, pól formularzy i tła stron; jedynie elementy
wizualne bez odpowiednika w Wordzie (wykresy, ikony, ozdobne pola) są
osadzane jako obrazy.

</details>

<details>
<summary><b>Czy GenOffice zbiera jakieś dane?</b></summary>

Oficjalne pakowane wersje domyślnie wysyłają ograniczoną telemetrię
użytkowania, a raportowanie można wyłączyć w każdej chwili w Ustawienia →
Ogólne. Telemetria nigdy nie wysyła treści dokumentów, nazw plików, ścieżek
plików, tożsamości konta ani adresów e-mail. Zobacz
[GenOffice Privacy](../../PRIVACY.md), aby poznać pełny wykaz zdarzeń i
ujawnianych danych.

</details>

## Bezpieczeństwo

Zobacz [SECURITY.md](../../SECURITY.md), aby poznać podejście do
bezpieczeństwa procesu (sandboxing renderera, walidacja IPC, ograniczanie
linków zewnętrznych) oraz modele zagrożeń dla treści generowanych przez AI.

## Podziękowania

GenOffice nie byłoby możliwe bez tych projektów open source:

- [Electron](https://www.electronjs.org/) — środowisko desktopowe dla każdej
  aplikacji.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — jądro
  interfejsu arkusza kalkulacyjnego, które rozwija Sheets.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause,
  dołączane przez
  [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer)) —
  silnik strumienia treści stojący za prawdziwą edycją tekstu i obrazów w
  PDF.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) i
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — renderowanie PDF i
  składanie dokumentów.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  edytory blokowe w Docs i Markdown.
- [CodeMirror](https://codemirror.net/) (MIT) — edytor źródła w HTML.
- [Konva](https://konvajs.org/) — renderowanie canvas dla Slides i wykresów
  Sheets.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — metryki
  składu tekstu dla złożonych systemów pisma.
- [calamine](https://github.com/tafia/calamine) i
  [IronCalc](https://github.com/ironcalc/IronCalc) — warstwy odczytu i
  obliczeń sidecara xlsx w Rust.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) — dekoder
  MicroType Express dla osadzonych czcionek PowerPoint, przeniesiony na
  TypeScript.
- [React](https://react.dev/) (MIT) — warstwa interfejsu użytkownika
  każdej aplikacji.
- [Mermaid](https://mermaid.js.org/) (MIT) i [KaTeX](https://katex.org/)
  (MIT) — diagramy i matematyka w Markdown i Docs.
- [opentype.js](https://opentype.js.org/) (MIT) — parsowanie czcionek do
  metryk i wyszukiwania glifów.
- [JSZip](https://stuk.github.io/jszip/) (MIT) i
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) — warstwy kontenera OOXML i XML.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) — zestaw ikon używany na wstążkach.
- [electron-updater](https://www.electron.build/) (MIT) — aktualizacje w
  aplikacji.
- Czcionki Liberation, Carlito, Caladea i Noto CJK (OFL/Apache-2.0) —
  dołączone czcionki dokumentów.

`npm run notices` regeneruje dołączone podsumowanie licencji stron trzecich
(`tools/gen-third-party-notices.mjs`); wszystkie zależności środowiska
wykonawczego są na licencjach MIT/Apache-2.0/BSD-3-Clause/OFL.

## Licencja

GenOffice jest licencjonowane na warunkach [Apache License 2.0](../../LICENSE), z
jednym wyjątkiem: katalog `ee/` jest zarezerwowany dla przyszłych modułów
enterprise i jest objęty [GenOffice Enterprise License](../../ee/LICENSE).

Nazwy i logo GenOffice oraz Genspark są znakami towarowymi Mainfunc, Inc.
Licencja Apache-2.0 nie daje prawa do ich używania (patrz sekcja 6); forki
powinny korzystać z własnej marki.

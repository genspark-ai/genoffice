# Szybki start: interfejs i podstawy

GenOffice to pakiet biurowy działający w całości na Twoim komputerze: jedno okno, jeden rząd kart, w którym mieści się sześć edytorów — Docs (edytor tekstu), Sheets (arkusze), Slides (prezentacje), PDF, Markdown i HTML. Pliki to prawdziwe pliki .docx / .xlsx / .pptx / .pdf, w pełni zgodne z Wordem, Excelem i PowerPointem. Nie jest potrzebne połączenie z siecią.

## Przegląd interfejsu

![Ekran strony głównej](img/home-screen.png)

Okno składa się z trzech części:

- **Pasek kart (u góry)**: każdy otwarty plik to jedna karta. Karta Strona główna, najbardziej z lewej, jest zawsze obecna i nie można jej zamknąć; pozostałe karty to Twoje dokumenty. Kliknij dwukrotnie kartę, aby zmienić nazwę pliku w miejscu.
- **Obszar zawartości**: edytor (albo strona główna) należący do aktywnej karty.
- **Pasek menu**: na macOS w systemowym pasku menu, na Windows/Linux u góry okna. Menu Plik/Edycja/Widok zmieniają się tak, aby pasowały do aktywnego edytora.

## Tworzenie dokumentu

Dowolny z tych sposobów:

- Kliknij kartę szybkiego tworzenia w sekcji [Szybki start](help://getting-started) na **stronie głównej** (AI Docs, AI Sheets, AI Slides, ...).
- Menu **Plik ▸ Nowy**: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML lub PDF.
- Przeciągnij plik na okno albo kliknij go dwukrotnie w menedżerze plików (jeśli GenOffice jest aplikacją domyślną).

Nowy dokument otwiera się bez tytułu; plik na dysku powstaje dopiero przy pierwszym zapisaniu.

## Otwieranie plików

- Menu **Plik ▸ Otwórz…** (⌘O/ctrl+O) otwiera systemowy selektor: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Kliknij dowolny element na liście **Ostatnie** na stronie głównej.
- Także `genoffice <plik>` wpisane w terminalu otwiera pliki.

## Model zapisywania

- **Zapis ręczny**: ⌘S/ctrl+S albo Plik ▸ Zapisz / Zapisz jako… Przy pierwszym zapisie zostaniesz poproszony o lokalizację i nazwę.
- **Autozapis** włącza się dopiero po tym, jak zapiszesz plik ręcznie przynajmniej raz — pliku PDF, który tylko czytałeś, nigdy nie da się zapisać po cichu w tle. Autozapis następuje krótko po zmianie zawartości.
- Zamknięcie karty z niezapisanymi zmianami najpierw pyta o Zapisz / Odrzuć / Anuluj.
- Każdy zapis jest atomowy (plik tymczasowy + zmiana nazwy), więc awaria zasilania nie zostawi uszkodzonego pliku.

## Popularne skróty

| Czynność              | macOS | Windows / Linux |
| --------------------- | ----- | --------------- |
| Nowy dokument         | ⌘N    | ctrl+N          |
| Otwórz                | ⌘O    | ctrl+O          |
| Zapisz                | ⌘S    | ctrl+S          |
| Zamknij kartę         | ⌘W    | ctrl+W          |
| Otwórz ten podręcznik | F1    | F1              |

Skróty używane wewnątrz każdego edytora (malarz formatów, znajdź i zamień, operacje na tabelach, ...) opisano w odpowiednich rozdziałach; Docs dodatkowo udostępnia wyszukiwalne okno skrótów klawiaturowych (patrz jego rozdział).

## Co dalej

- Gdzie są pliki: [Ekran strony głównej](help://home-screen).
- Zarządzanie wieloma otwartymi plikami: [Karty i zarządzanie oknami](help://tabs-and-windows).
- Zlecenie pracy sztucznej inteligencji: [Panel asystenta AI](help://ai-panel).
- Język, motyw, aplikacje domyślne: [Ustawienia, język, motyw i integracje MCP](help://settings-integrations).

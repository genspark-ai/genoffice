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
| Zwiń Wstążkę          | ⌥⌘R   | Ctrl+F1         |

**Zwinięcie wstążki** działa w każdym edytorze. Wiersz kart pozostaje na miejscu, a pasek poleceń pod nim znika; wybrana karta jest zarazem przyciskiem zwijania, więc dopóki wstążka jest zwinięta, żadna karta nie jest wybrana, a naciśnięcie dowolnej karty przywraca pasek. Podwójne kliknięcie karty robi dokładnie to samo. Sposób, w jaki ją zostawili, jest zapamiętywany osobno dla każdego edytora.

Skróty używane wewnątrz każdego edytora (malarz formatów, znajdź i zamień, operacje na tabelach, ...) opisano w odpowiednich rozdziałach; Docs dodatkowo udostępnia wyszukiwalne okno skrótów klawiaturowych (**⌘/**) (patrz jego rozdział).

## Skróty Option+Command

Option+Command to warstwa, którą Word rezerwuje na strukturyzowane skoki, a GenOffice wypełnia ją w ten sam sposób. Docs zajmuje większość z nich, Sheets dwa własne na potrzeby zgodności z Excelem, a jeden skrót działa wszędzie.

**Docs**

| Skrót (macOS)   | Co robi               | Windows / Linux    |
| --------------- | --------------------- | ------------------ |
| ⌥⌘1 / ⌥⌘2 / ⌥⌘3 | Nagłówek 1 / 2 / 3    | Ctrl+Alt+1 / 2 / 3 |
| ⌥⌘0             | Normalny              | Ctrl+Alt+0         |
| ⌥⌘M             | Akapit                | Ctrl+Alt+M         |
| ⌥⌘A             | Nowy komentarz        | Ctrl+Alt+A         |
| ⌥⌘F             | Wstaw przypis dolny   | Ctrl+Alt+F         |
| ⌥⌘E             | Wstaw przypis końcowy | Ctrl+Alt+D         |
| ⌥⌘G             | Przejdź do            | Ctrl+G             |

Dwa z nich działają inaczej w Windows, z tego samego powodu, dla którego Word je rozdziela. **macOS zajmuje ⌥⌘D** — pokazuje i ukrywa Dock —, więc przypis końcowy to ⌥⌘E na Macu, a wszędzie indziej Ctrl+Alt+D. A **Przejdź do** rezygnuje z Alt: Ctrl+G, tam gdzie skrót na Macu zabiera go ze sobą.

**Sheets**, gdy siatka ma fokus

| Skrót (macOS) | Co robi              | Windows / Linux |
| ------------- | -------------------- | --------------- |
| ⌥⌘0           | Krawędzie zewnętrzne | Ctrl+Shift+7    |
| ⌥⌘−           | Brak krawędzi        | Ctrl+Shift+−    |

Windows to nie przepisanie pary z Maca. Excel na Macu daje Sheets **oba** — ⌘⇧7 i ⌥⌘0 to dwa klawisze do tych samych krawędzi zewnętrznych —, więc w Windows polecenie zachowuje przypisane już miejsce Ctrl+Shift, a warstwa Option po prostu nie ma.

Zwróć uwagę, że **⌥⌘0 oznacza Normalny w Docs, a Krawędzie zewnętrzne w Sheets**. Nigdy nie pojawiają się w tym samym edytorze, więc w praktyce nic się nie koliduje, ale ⌥⌘0 jest już zajęty i niedostępny jako skrót globalny.

**Każdy edytor**: **⌥⌘R / Ctrl+F1** zwija wstążkę, jak opisano powyżej.

To zostawia ⌥⌘D wolny dla GenOffice na macOS, gdyby jakieś przyszłe polecenie go potrzebowało.

## Co dalej

- Gdzie są pliki: [Ekran strony głównej](help://home-screen).
- Zarządzanie wieloma otwartymi plikami: [Karty i zarządzanie oknami](help://tabs-and-windows).
- Zlecenie pracy sztucznej inteligencji: [Panel asystenta AI](help://ai-panel).
- Język, motyw, aplikacje domyślne: [Ustawienia, język, motyw i integracje MCP](help://settings-integrations).

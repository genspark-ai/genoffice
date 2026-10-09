# Ekran strony głównej: gdzie mieszkają Twoje pliki

Strona główna to punkt wyjścia GenOffice: pasek nawigacyjny po lewej, listy plików i karty szybkiego tworzenia po prawej.

![Ekran strony głównej](img/home-screen.png)

## Nawigacja na pasku bocznym

- **Ostatnie**: pliki, które ostatnio otwierałeś. Każdy wiersz ma na sobie czas — dzisiaj, wczoraj albo datę.
- **Ulubione**: pliki oznaczone gwiazdką. Najedź na wiersz pliku i kliknij gwiazdkę, aby dodać lub usunąć plik.
- **Podręcznik użytkownika**: otwiera ten podręcznik.
- **Genspark Projects**: po zalogowaniu na konto Genspark pokazuje projekty utworzone w sieci za pomocą Genspark AI; kliknięcie projektu otwiera go w przeglądarce do dalszej edycji. Dostępne są wyszukiwanie, sortowanie według czasu, odświeżanie i wczytanie kolejnych.
- **Foldery**: katalogi dodane do paska bocznego przyciskiem **Dodaj folder…** albo przeciągnięte tam z menedżera plików. Każdy staje się katalogiem głównym, który można otworzyć, utworzyć w nim podfoldery, zmienić jego nazwę i usunąć; ten, który przestanie być dostępny, jest oznaczany jako niedostępny i można go zdjąć z listy. **Nowy folder** tworzy kolejny.

Nie ma tu pozycji Kosz. Usunięte pliki trafiają do kosza systemowego, a ich przywrócenie to zadanie systemu operacyjnego.

## Lista plików

Każdy wiersz pokazuje ikonę, nazwę pliku, czas modyfikacji i inne dane. **⋯ menu** w wierszu oferuje:

- **Otwórz**, a także **Pokaż w folderze**, aby zlokalizować plik w menedżerze plików.
- **Kopiuj ścieżkę**.
- **Przenieś do folderu…**: otwiera wybór folderu i naprawdę przenosi plik; jeśli w miejscu docelowym jest już plik o tej nazwie, możesz go pominąć, nadpisać albo zmienić jego nazwę.
- **Zmień nazwę**: zmiana w miejscu, rozszerzenie zachowywane automatycznie.
- **Dodaj do ulubionych / Usuń z ulubionych** — gwiazdki pozostają po restarcie i towarzyszą plikowi także po jego przemianowaniu.
- **Utwórz kopię**: tworzy kopię w tym samym folderze.
- **Usuń**: przenosi plik do kosza systemowego — to nie jest trwałe usunięcie.
- **Usuń z listy**, w głównym widoku **Ostatnie**, aby zniknęła pozycja bez ruszania pliku.

### Wiele plików naraz

Zaznacz pole wyboru w wierszu lub kliknij ⌘/ctrl, aby zbudować zaznaczenie; pole w nagłówku zaznacza wszystko, co jest teraz na liście, a pasek nad listą podaje (**Wybrano: {n}**), ile plików jest zaznaczonych, i oferuje **Przenieś do folderu…** oraz **Usuń pliki** dla całego zestawu. Zaznaczenie wielokrotne możesz też przeciągnąć na folder w pasku bocznym.

## Wyszukiwanie

Pole wyszukiwania u góry filtruje naraz dwie rzeczy:

- **Nazwy plików**: szybkie filtrowanie po nazwie.
- **Zawartość plików**: GenOffice indeksuje Twoje pliki w tle (tekst w docx/xlsx/pptx/pdf/md/html), więc wyszukiwanie po treści znajduje również pliki. Zakres i przełączniki ustawia się w ustawieniach wyszukiwania.

## Karty szybkiego startu

Karty nad listami tworzą nowy dokument jednym krokiem. Kliknięcie karty tworzy plik danego typu i otwiera jego edytor — możesz od razu pisać albo pozwolić, aby AI przygotował wstęp (każdy edytor ma **przycisk AI** na wstążce oraz **Zapytaj AI** w menu kontekstowym zaznaczenia).

Nowy plik trafia do folderu aktualnie zaznaczonego na pasku bocznym; jeśli nic nie jest zaznaczone, trafia do folderu domyślnego.

Co robi każda karta:

- **AI Docs** (.docx): pusty dokument tekstowy w edytorze Docs. Plik zostaje zapisany na dysku dopiero przy **pierwszym zapisaniu**; nowe dokumenty otwierają się z rozwiniętym panelem AI (wyłączasz to w Ustawienia → „Otwieraj panel AI w nowych dokumentach”).
- **AI Sheets** (.xlsx): pusty arkusz w edytorze Sheets. Dopóki nie zapiszesz, na dysku nie ma żadnego pliku — nazwa jest zarezerwowana na pierwsze zapisanie; po pierwszym wygenerowaniu przez AI plik można też automatycznie zmienić na podstawie jego zawartości.
- **AI Slides** (.pptx): pusta prezentacja w edytorze Slides.
- **AI Markdown** (.md): pusty dokument Markdown w edytorze Markdown.
- **AI HTML** (.html): pusta strona internetowa w edytorze HTML.
- **AI PDF** (.pdf): inaczej niż pozostałe — **natychmiast** tworzy prawdziwy pusty plik PDF o jednej stronie w folderze docelowym i otwiera go jako zwykły plik (edytor PDF pracuje na prawdziwych plikach). Dobra do nanoszenia adnotacji, zaciemniania i dodawania tekstu; przy pierwszym zapisaniu plik można automatycznie zmienić na podstawie jego zawartości.
- **Otwórz plik lokalny**: systemowy selektor plików dla programów Word (.docx/.doc), Excel (.xlsx/.xlsm/.xls/.csv/.tsv), PowerPoint (.pptx/.ppt), PDF, Markdown (.md/.markdown) i stron internetowych (.html/.htm). Można zaznaczyć kilka plików; każdy dostaje własną kartę.

> Wskazówka: Plik ▸ Nowy na pasku menu tworzy te same typy dokumentów (⌘N/Ctrl+N domyślnie tworzy dokument tekstowy); przeciągnięcie pliku do okna otwiera go.

## Projekty w chmurze (Genspark Projects)

- Przy pierwszym użyciu trzeba zalogować się na konto Genspark (przepływ z kodem urządzenia: GenOffice pokazuje kod, a Ty kończysz logowanie w przeglądarce).
- Lista projektów synchronizuje się z siecią; Otwórz w przeglądarce przenosi tam, gdzie można kontynuować.
- Brak logowania nie wpływa na żadną funkcję lokalną.

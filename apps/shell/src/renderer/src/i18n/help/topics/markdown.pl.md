# Edytor Markdown

Edytor Markdown otwiera pliki .md / .markdown w układzie źródło podgląd wyrenderowany.

- **Otwórz**: ze strony głównej lub Plik ▸ Otwórz…; działa też z wiersza poleceń.
- **Edycja**: edycja w czystym tekście; rozszerzenia GFM (tabele, listy zadań, przekreślenie, automatyczne odnośniki) są wyrenderowywane w podglądzie.
- **Podgląd**: na żywo; zasoby względne, takie jak obrazy, rozwiązywane są obok dokumentu.
- **Zapisywanie**: wierne bajt w bajt — znacznik BOM, znaki CRLF i obecność pustej linii na końcu są zachowywane; zapis bez zmian nie przepisuje pliku.
- **Znajdź i zamień**: ctrl+F przeszukuje źródło; Zastąp wszystko zapisuje wynik.
- **AI**: gotowe przyciski pozwalają asystentowi przepisać, rozbudować lub przetłumaczyć dokument.

## Pasek narzędzi

Jeden rząd przycisków nad edytorem (najedź, aby zobaczyć opisy):

![Pasek narzędzi Markdown](img/md-toolbar.png)

- **Plik i historia**: Zapisz, Zapisz jako…, Cofnij, Ponów, Znajdź; przełącznik **Autozapisu** po prawej zapisuje zmiany na dysku w określonych odstępach czasu.
- **Przycisk AI**: otwiera panel AI; obok są ustawienia wstępne przepisz / rozwiń / przetłumacz.
- **Styl akapitu** (lista rozwijana): przełączanie między tekstem głównym a poziomami nagłówków.
- **Formatowanie w tekście**: **pogrubienie**, _pochylenie_, ~~przekreślenie~~, `kod w tekście`, odnośnik.
- **Listy**: punktowana, numerowana, lista zadań.
- **Wstaw**: tabela, obraz, pozioma linia.
- **Właściwości**: wstawia lub przechodzi do bloku YAML front matter na początku pliku.
- **Konspekt**: przechodzenie według hierarchii nagłówków.
- **Sprawdzanie pisowni**: włącza lub wyłącza sprawdzanie pisowni dla tego dokumentu.

Trzy szybkie przykłady:

- **Nagłówek**: umieść kursor w wierszu ▸ lista rozwijana stylu akapitu ▸ „Nagłówek 1”.
- **Tabela**: kliknij **Wstaw tabelę** ▸ przeciągnij, aby wybrać liczbę wierszy i kolumn ▸ wpisz zawartość komórek; podgląd renderuje ją natychmiast.
- **Lista zadań**: zaznacz kilka wierszy ▸ kliknij **Lista zadań** ▸ każdy wiersz staje się `- [ ]`, w podglądzie pokazywany jako pola wyboru.

## Eksportowanie

Menu Plik — wszystko lokalne i wszystko pyta, gdzie umieścić wynik:

- **Eksportuj jako Word…** i **Eksportuj jako PDF…** zapisują prawdziwy plik .docx lub .pdf.
- **Eksportuj jako obrazy…** zapisuje po jednym PNG na stronę w katalogu, który wybierzesz.
- **Konwertuj i otwórz w Docs** konwertuje do .docx i otwiera go we wbudowanej karcie Docs tutaj w aplikacji — to nie jest przekazanie czegokolwiek do chmury, a przekonwertowana kopia leży w folderze pamięci podręcznej, który jest czyszczony po około tygodniu.

## Widok źródła

Wstążka zawiera przełącznik **Źródło** (z lokalizacją razem z aplikacją). Włącz go, a edytor zostanie zastąpiony surowym Markdownem: dokładnie ten tekst, który zapisuje zapisywanie, nic nie jest ładniejsze, nic nie jest normalizowane w tle.

- **Edycja wierna bajtom.** Zapis z widoku źródła daje te same bajty co zapis z edytora — BOM, CRLF i obecność końcowego znaku nowej linii zostają zachowane.
- **To ten sam dokument.** Przełączaj się swobodnie; źródło to własny tekst edytora, a nie kopia wymagająca scalenia.
- **Pasek narzędzi formatowania je niedostępny** podczas jej otwarcia, bo większość tych przycisków wstawia konstrukcje edytora, które mają sens tylko po wyrenderowaniu. Wraca po zamknięciu widoku.
- **Pliki JSON i inne w trybie źródłowym** otwierają się tutaj bezpośrednio: nie ma czego renderować, więc źródło _jest_ dokumentem.

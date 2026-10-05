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

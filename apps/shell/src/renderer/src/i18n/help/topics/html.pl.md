# Edytor HTML

Edytor HTML otwiera pliki .html / .htm w dwóch trybach: **podgląd** (wyrenderowana strona) i **źródło**.

- **Podgląd**: rzeczywiste renderowanie; względne arkusze stylów i obrazy wczytują się obok pliku.
- **Inspektor podglądu**: kliknij, aby zaznaczyć element, kliknij dwukrotnie, aby edytować jego tekst w miejscu, usuń z paska narzędzi oraz Zapytaj AI o zaznaczenie.
- **Tryb źródła**: edycja HTML; ctrl+F do wyszukiwania, Zastąp wszystko zapisuje przepisany znacznik.
- **Zapisywanie**: wierne bajt w bajt (BOM/CRLF/końcowa nowa linia zachowane); zapis bez zmian nic nie przepisuje.
- **Powiększenie**: ctrl+kółko / gest szczypania zmienia skalę podglądu; ctrl+Z wewnątrz podglądu cofa ostatnią zmianę.

## Pasek narzędzi

Kliknij dowolny element w podglądzie, a nad nim pojawi się pasek narzędzi:

![Pływający pasek narzędzi nad zaznaczonym elementem](img/html-toolbar.png)

- **Plik i historia**: Zapisz, Zapisz jako…, Cofnij, Ponów, Znajdź; przełącznik **Autozapisu** zapisuje zmiany w określonych odstępach czasu.
- Przełącznik **Podgląd / Źródło**; **Prezentuj** pokazuje stronę na pełnym ekranie.
- **Formatowanie**: pogrubienie, pochylenie, zwiększanie i zmniejszanie rozmiaru czcionki; **panel stylów** dla zaznaczonego elementu (kolory i inne).
- **Wstaw**: nagłówek, akapit, tabela, obraz (przez odnośnik), inne.
- **Operacje na obrazie** (z zaznaczonym obrazem): przytnij, **usuń tło**, zastąp, zablokuj proporcje.
- **Operacje na elemencie** (z elementem zaznaczonym w inspektorze podglądu): usuń, zduplikuj, przesuń w górę/w dół.
- **Przycisk AI**: otwiera panel AI; możesz pytać o zaznaczony element.

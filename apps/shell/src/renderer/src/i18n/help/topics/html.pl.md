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

## Eksportowanie

Menu Plik — wszystko lokalne i wszystko pyta, gdzie umieścić wynik:

- **Eksportuj jako Word…** i **Eksportuj jako PDF…** zapisują prawdziwy plik .docx lub .pdf.
- **Eksportuj jako pojedynczy plik HTML…** zapisuje jeden .html z osadzonymi obrazami. Nie nadpisuje pliku, który masz teraz otwarty, i mówi, ile obrazów nie udało się osadzić.

## Wstaw szkielet

Dla pustej strony **Wstaw ▸ Wstaw szkielet** zapisuje minimalny dokument w trybie standardowym:

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <title></title>
  </head>
  <body></body>
</html>
```

Każda część jest tam z jakiegoś powodu — dlatego to polecenie, a nie coś, co się wpisuje:

- **doctype**, bo inaczej podgląd działa w trybie quirks, w którym rozmiary ramek i układ tabel podlegają innym regułom, niż się spodziewasz;
- **`lang`**, bo inaczej czytnik ekranu nie ma języka, w którym miałby czytać stronę, a przeglądarka dobiera czcionkę i sprawdzanie pisowni do złego języka;
- **`charset`**, bo inaczej strona z tekstem niełacińskim może zostać zdekodowana jako mojibake.

Meta viewport celowo nie ma: to renderuje się w panelu pulpitu bez mobilnego viewportu, na który miałby wpływ.

`lang` podąża za językiem interfejsu aplikacji, więc wstawiany szkielet jest tym, pod który twoje narzędzia są już skonfigurowane. Potem możesz go swobodnie edytować.

Ta pozycja pojawia się tylko w trybie edycji i tylko, gdy dokument jest pusty — gdy jest treść, nie ma czegoś, _w co_ wstawić szkielet.

# Sheets: arkusze

Sheets to edytor podobny do Excela; obliczenia obsługuje osobny proces silnika napisanego w Rust (awaria w nim nie wywróci aplikacji). Otwiera i zapisuje prawdziwe pliki .xlsx; pliki .csv i .tsv otwierają się jako tabele.

## Interfejs

- **Wstążka**: osiem kart, opisanych po kolei poniżej.
- **Pasek formuły**: pokazuje i pozwala edytować formułę bieżącej komórki; obsługiwane są typowe funkcje.
- **Karty arkuszy** (na dole): dodaj / zmień nazwę / usuń / przenieś arkusz.
- **Edycja komórki**: kliknij dwukrotnie albo po prostu pisz; Enter zatwierdza i przechodzi w dół, Tab w prawo, Esc anuluje (nawyki z Excela).
- **Skróty**: zgodne z rodziną Excela (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, ...).

## Karty wstążki

- **Narzędzia główne**: czcionka, wypełnienie, obramowania, formaty liczbowe (waluta/procent/separator tysięcy, zwiększanie i zmniejszanie miejsc po przecinku), wyrównanie, scalanie, wstawianie wierszy i kolumn oraz ich rozmiary, formatowanie warunkowe, formatuj jako tabela, style komórek, schowek i malarz formatów, sortowanie i filtr.
- **Wstawianie**: obrazy, kształty, pola tekstowe, hiperłącza, komentarze, pole wyboru, nagłówek i stopka, symbole, równania; **dziewięć typów wykresów** (kolumnowy, słupkowy, liniowy, warstwowy, kołowy, punktowy, radarowy, pierścieniowy oraz kombinowany kolumnowo-liniowy) plus **Polecane wykresy**, które układają typy dla zaznaczenia i pokazują ich podgląd; **Wykres przestawny** z komórki, w której Pan/Pani jest; **Wykresy przebiegu w czasie** (liniowy, kolumnowy, zwycięzca/przegrana); **Fragmentator** i **Oś czasu** do filtrowania tabeli przestawnej.
  - Wykresy, wykresy przestawne i wykresy przebiegu w czasie są rzeczywistymi obiektami w zapisanym skoroszycie.
  - **Fragmentator** i **Oś czasu** to elementy sterujące sesji: filtrowanie, które wykonują, jest zapisywane, a Excel pokazuje tę samą przefiltrowaną tabelę przestawną, ale sam przycisk fragmentatora nie jest częścią pliku.
- **Układ strony**: kolory i czcionki motywu, przełączniki drukowania linii siatki i nagłówków, podgląd podziału na strony.
- **Formuły**: Suma automatyczna i wstawianie funkcji, definiowanie nazw (także z zaznaczenia), śledzenie poprzedników i zależności, okno obserwacji, przeliczanie arkusza lub skoroszytu.
- **Dane**: sortowanie i filtrowanie (także filtr zaawansowany i wyczyszczenie filtra), tekst na kolumny, scalanie skoroszytów, odśwież wszystko.
- **Recenzja**: przeglądanie komentarzy (pokaż, poprzedni/następny), tłumaczenie oraz grupa **Ochrona** — **Chroń arkusz**, **Chroń skoroszyt** i **Zezwalaj na edytowanie zakresów**.
  - Ochrona jest zapisywana w pliku .xlsx, a to, co nakłada ta aplikacja, nie ma hasła, więc ten sam przycisk zmienia się w **Nie chroń…** i cofa ją. Ochrony nałożonej przez inny program, który _ma_ hasło, nie da się stąd usunąć.
  - Oba przyciski ochrony działają natychmiast — nie ma okna dialogowego, które można by anulować, tylko uwaga na pasku stanu, że zostanie zapisane przy zapisie.
  - **Zezwalaj na edytowanie zakresów** zaznacza komórki, które pozostają edytowalne, podczas gdy reszta arkusza jest zablokowana.
- **Widok**: przełączniki linii siatki, **pasku formuły**, nagłówków i wyróżniania aktywnego wiersza i kolumny; powiększenie; **Normalny** i **Podgląd podziału na strony**. Linie siatki i nagłówki są zapisywane wraz z arkuszem; wyróżnianie to własne ustawienie.
- **Projekt wykresu**: pojawia się po zaznaczeniu wykresu — typ wykresu, style i kolory, edycja zakresu danych.

Karta Dane, przycisk po przycisku (od lewej do prawej na obrazku):

![Karta Dane](img/sheets-data.png)

- **Tabela przestawna**: buduje tabelę przestawną z bieżącego zakresu; przeciągnij pola, aby agregować.
- **Odśwież**: przelicza dane bieżącej tabeli przestawnej.
- **Z tekstu/CSV**: importuje plik .csv/.txt jako nowy arkusz, dzieląc go według separatora.
- **Scal skoroszyty**: pobiera arkusze z innych plików .xlsx do tego pliku.
- **Odśwież wszystko**: przelicza wszystkie tabele przestawne i zewnętrzne źródła danych.
- **Sortuj** (lista rozwijana): rosnąco / malejąco / sortowanie niestandardowe (reguły dla wielu kolumn).
- **Filtruj**: dodaje listy ▼ do wiersza nagłówków; zaznacz wartości, które chcesz zachować.
- Małe przyciski ułożone obok siebie: **Wyczyść** (przywraca wszystkie wiersze), **Zastosuj ponownie** (uruchamia bieżący filtr jeszcze raz), **Zaawansowane** (filtruje według zakresu kryteriów).
- **Tekst na kolumny** (lista rozwijana): dzieli jedną kolumnę na kilka według separatora lub stałej szerokości.
- **Wypełnienie szybkie**: podaj jeden przykład, a reszta kolumny wypełni się według niego (ctrl+E).
- **Usuń duplikaty**: odrzuca zduplikowane wiersze na podstawie zaznaczonych kolumn.
- **Sprawdzanie poprawności** (lista rozwijana): reguły wprowadzania danych dla zaznaczenia (listy, zakresy liczb...).
- **Scalaj**: agreguje kilka zakresów w jedno miejsce według kategorii.
- **Analiza what-if** (lista rozwijana): szukanie celu — rozwiązuje jedną komórkę wejściową tak, aby komórka z formułą osiągnęła wartość docelową.
- **Grupuj / Anuluj grupowanie** (lista rozwijana): grupy wierszy lub kolumn ze zwijaniem i rozwijaniem.
- **Podsumowanie**: wstawia wiersze sum częściowych dla każdej kategorii.

Karta Formuły, przycisk po przycisku:

![Karta Formuły](img/sheets-formulas.png)

- **Wstaw funkcję** (fx): wyszukuje funkcje wraz z kreatorem argumentów.
- **Suma automatyczna** (lista rozwijana): SUMA jednym kliknięciem, a także średnia/ile/ maksimum/minimum.
- **Ostatnio używane / Finansowe / Logiczne / Tekstowe / Data i godzina / Wyszukiwanie i odwołania / Matematyczne i trygonometryczne / inne**: przeglądaj i wstawiaj funkcje według kategorii.
- **Menedżer nazw**: wyświetlanie, tworzenie i usuwanie nazwanych zakresów.
- **Zdefiniuj nazwę** (lista rozwijana): nadaje zaznaczeniu nazwę; **Użyj w formule** wstawia istniejącą nazwę; **Utwórz z zaznaczenia** nadaje nazwy zakresom na podstawie wiersza lub kolumny nagłówkowej.
- **Śledź poprzedniki / Śledź zależności**: niebieskie strzałki pokazujące, skąd pochodzą dane formuły i dokąd trafiają; **Usuń strzałki** je usuwa.
- **Pokaż formuły**: komórki wyświetlają samą formułę zamiast wyniku.
- **Sprawdzanie błędów**: wyszukuje i objaśnia błędy formuł.
- **Okno obserwacji**: przypina komórki, którymi się interesujesz, i śledzi ich bieżące wartości.
- **Opcje obliczania** (lista rozwijana): przeliczanie automatyczne lub ręczne; w trybie ręcznym **Oblicz teraz / Oblicz arkusz** uruchamia je na żądanie.

## Liczby i formatowanie

- Formaty liczbowe: ogólny, liczbowy, walutowy, procentowy, data/godzina, ułamkowy, naukowy i inne.
- Wyrównanie, zawijanie tekstu, scalone komórki, obramowania i wypełnienia.
- Wysokości wierszy i szerokości kolumn przez przeciąganie; kliknięcie dwukrotnie krawędzi dopasowuje rozmiar automatycznie.

## Dane

**Sortowanie i filtrowanie** (na przykład malejąco według jednej kolumny):

1. Kliknij **dowolną komórkę w tej kolumnie** (nie trzeba zaznaczać całej kolumny).
2. Karta Narzędzia główne ▸ **Sortuj i filtruj** ▸ **Malejąco**; całe wiersze przestawiają się razem (cały obszar jest sortowany jako jedno).
3. Dla własnych reguł (wiele kolumn, według koloru): ta sama ścieżka, **Sortowanie niestandardowe**.
4. Filtr: zaznacz wiersz nagłówków i kliknij **Sortuj i filtruj ▸ Filtruj** — każdy nagłówek dostaje listę ▼, w której zaznaczasz wartości do zachowania; wyczyszczenie filtra przywraca wszystko.

- Sortowanie i filtrowanie.
- Zablokowanie okien.
- .csv / .tsv: otwierają się od razu jako tabela (tsv rozdzielony tabulatorami jest parsowany jako jeden); zapisywanie zapisuje plik w oryginalnym formacie.

## Menu kontekstowe

- **W siatce**: własne menu edytora (Univer) — wytnij/kopiuj/wklej, wstawianie i usuwanie wierszy i kolumn, ukrywanie, scalanie komórek, blokowanie okien i inne codzienne pozycje.
- **Na dolnym pasku stanu**: wybierz, które statystyki ma pokazywać pasek stanu (średnia / liczba / suma, ...); wybór zostaje zapamiętany.
- **Na karcie arkusza na dole**: dodaj / zmień nazwę / usuń / pokoloruj / ukryj arkusze (menu kart Univer).
- Menu kontekstowe górnego paska kart opisano w [Karty i zarządzanie oknami](help://tabs-and-windows).

## AI

- Boczny panel AI: zaznacz zakres i wydaj polecenie w zwykłym języku (zmień formatowanie, wygeneruj dane, napisz formuły).
- Dołącz pliki do polecenia przyciskiem 📎 albo przeciągnij je na panel; towarzyszą pytaniu, a obrazy wracają jako miniatury.
- Odpowiedź może wskazać komórkę — kliknięcie odwołania przenosi siatkę w to miejsce.
- Zmiany wprowadzone przez AI można cofnąć z panelu.

## Zapisywanie i eksport

- Zapis .xlsx (z zachowaniem formuł i formatów); Zapisz jako…; eksport do PDF stosuje podział na strony używany przy drukowaniu.
- Autozapis stosuje się do reguły globalnej (włącza się po pierwszym zapisaniu ręcznym).

## Stabilność

- Pomocniczy proces obliczeniowy w Rust jest odizolowany od interfejsu: jeśli wyjątkowe dane go zatrzymają, zobaczysz komunikat i próbę odzyskania sesji — a nie awarię aplikacji.

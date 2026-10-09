# PDF: czytanie, adnotowanie i zaciemnianie

Edytor PDF ma pięć kart na wstążce: **Narzędzia główne / Adnotacje / Edycja / Strony / Widok**. Czyta i zapisuje: tekst można edytować, treść zaciemniać i podpisywać, a formularze wypełniać.

## Czytanie i nawigacja

- Lewy pasek boczny: **miniatury** (kliknij, aby przejść, widoczny zakres wyróżniony) albo **spis treści** (zakładki, jeśli są).
- Powiększenie: regulator współczynnika w prawym dolnym rogu; ctrl+kółko zmienia powiększenie stopniowo.
- Obracanie: dla jednej strony lub wszystkich z menu Strony; obroty są zapisywane w pliku przy zapisie.
- Wyszukiwanie: ctrl+F na pełnym tekście z podświetleniem wszystkich trafień.
- Zaszyfrowane pliki PDF: otwiera je monit o haśle (we własnym małym oknie); hasło jest używane tylko w tej sesji.

## Zaznaczanie tekstu i oznaczanie (Adnotacje)

Spróbuj na dowolnym akapicie:

1. **Przeciągnij mysz przez zdanie** — po puszczeniu nad tekstem pojawia się pasek adnotacji:

![Pasek adnotacji po zaznaczeniu tekstu](img/pdf-highlight.png)

2. Wybierz **wyróżnienie** (żółta próbka otwiera paletę kolorów), **podkreślenie** albo **przekreślenie**; **Zapytaj AI** wysyła zaznaczenie razem z Twoim pytaniem do panelu AI.
3. Aby cofnąć adnotację, zaznacz przeciągnięciem ten sam fragment i kliknij aktywny przycisk na pasku (przełącznik w stylu Worda), albo zaznacz go i naciśnij Usuń.

Oczekiwane działanie:

- Przeciągnięcie po tekście wyświetla pasek: **wyróżnienie / podkreślenie / przekreślenie / kopiuj / Zapytaj AI**.
- Kolory pochodzą z palety; **nałożenie tego samego oznaczenia na już oznaczony zakres je usuwa** (przełącznik w stylu Worda).
- Oznaczenia zapisane już w pliku można zaznaczyć i usunąć (⋯ menu lub Usuń).
- **Uwaga**: gdy aktywne jest narzędzie rysowania, warstwa tekstu nie jest dostępna do zaznaczenia — narzędzie wyłącza się samo po każdym umieszczeniu, więc przed następną czynnością znów jesteś w trybie zaznaczania.

## Narzędzia rysowania (Adnotacje)

Sześć narzędzi: **tusz, prostokąt, elipsa, strzałka, notatka**, a ponadto **ramka zaciemniania** w karcie Adnotacje.

- Każde narzędzie jest przełącznikiem: kliknij, aby je włączyć; **wyłącza się samo, gdy tylko zostanie umieszczona figura** (aby kontynuować, kliknij narzędzie ponownie); kliknięcie już włączonego narzędzia również je wyłącza.
- Tusz podąża za grubością pociągnięcia; prostokąt, elipsę i strzałkę rysuje się przez przeciągnięcie; kolory pochodzą z palety rysowania.
- Umieszczone figury można zaznaczyć, usunąć, przeciągnąć oraz (prostokąt i elipsę) przeskalować.
- **Zaciemnianie, cały przepływ** (na przykład ukrycie wiersza tekstu):

  1. Karta Adnotacje ▸ kliknij **Zaczernij obszar** (narzędzie się włącza).
  2. **Przeciągnij ramkę na treść** — zostanie zakryta znacznikiem kreskowanym, a na pasku narzędzi pojawią się przyciski **wyczyść oznaczenia / zastosuj zaczernienia**:

  ![Strona po zaznaczeniu zaciemnienia](img/pdf-redact.png)

  3. Kliknij **Zastosuj zaczernienia** i potwierdź — powstanie kopia robocza, w której zakryty tekst i obrazy są fizycznie usunięte (a nie tylko przykryte), a operacji nie można cofnąć; oryginalny dokument pozostaje nietknięty.

  Pomyliłeś się? Przycisk wyczyszczenia oznaczeń kasuje bieżące oznaczenia, więc możesz narysować od nowa.

## Przypięte notatki i wątki komentarzy

- **Narzędzie notatki** umieszcza pinezkę i otwiera kartę przy marginesie na tekst (imię autora można ustawić); po potwierdzeniu zapisuje się jako standardowa adnotacja tekstowa PDF.
- Kliknij pinezkę, aby otworzyć wątek: **odpowiedz** (płaskie wątki w stylu WPS/Acrobata), **edytuj** własny komentarz, **usuń** pojedynczy komentarz lub cały wątek.
- Trwające zmiany pozostają aktywne, dopóki zapis nie wpisze nowego tekstu do tej samej adnotacji w pliku, zachowując łańcuch odpowiedzi.

## Edycja zawartości PDF (Edycja)

- **Edytuj tekst**: kliknij tekst, aby edytować go blokami (silnik pdfium; dopasowanie krojów najlepsze, jak to możliwe).
- **Wstaw tekst**: wstawia przeszukiwalny tekst o wybranym kroju, rozmiarze i kolorze.
- **Wstaw obraz / pieczęć**.
- Formularze: pola AcroForm wypełnia się bezpośrednio; wartości są zapisywane przy zapisie.

## Podpisy

- **Podpis odręczny**: narysuj go; można go powiązać z polem podpisu w formularzu.
- **Podpis obrazkowy**: umieść obraz jako podpis.
- Zapisane podpisy można wykorzystywać ponownie.

## Operacje na stronach (Strony)

- **Obróć / usuń / zmień kolejność**: przeciągaj miniatury, aby zmienić kolejność; usuwanie wymaga potwierdzenia.
- **Importuj strony**: wciągnij strony z innego pliku PDF. **Wstaw pustą stronę** dodaje pustą stronę.
- **Zamień strony** podmienia cały zakres na strony skąd indziej; **Przytnij strony** przycina krawędzie, z opcją zastosowania tego do wszystkich stron.
- **Rozmiar strony** skaluje wszystkie strony do jednego rozmiaru papieru; **Odwróć kolejność** odwraca dokument od końca do początku.
- **Wyodrębnij strony**: eksportuje wybrane strony do nowego pliku PDF.
- **Podziel PDF**: dwie postacie — podział według zakresów na kilka plików albo pocięcie każdej strony na siatkę mniejszych stron.
- **Scal PDF**: dwie postacie — dołączenie innych plików PDF albo połączenie kilku stron na jednym arkuszu. Rozmiary są sumowane **zanim cokolwiek zostanie odczytane**, a łączna wartość powyżej **1 GiB jest odrzucana** z czytelnym komunikatem błędu (aby ograniczyć zużycie pamięci).
- Zmiany na poziomie strony są zapisywane przy następnym zapisie; Zapisz jako… pozostawia oryginał nietknięty.

## Eksport i drukowanie

- **Eksportuj jako Word… / PowerPoint… / Excel…** w menu Plik albo te same trzy z **Konwertuj PDF** na wstążce — wszystko lokalnie, bez wysyłania na serwer. Plik .pptx wychodzi ze slajdem na stronę, a .xlsx z arkuszem na stronę. Każdy z nich pyta, gdzie zapisać.
- **Drukuj**: bieżąca kolejność i obroty przez okno systemowe; obsługiwane są zakresy stron.

## Zapisywanie

- Zwykłe zapisywanie, ręczne lub automatyczne, zapisuje adnotacje i edycje z powrotem w pliku (atomowo).
- **Zaciemnianie przechodzi przez własny przepływ aplikowania** i tworzy kopię, pozostawiając oryginał nietknięty, dzięki czemu wrażliwa treść nie pozostaje w pliku źródłowym.

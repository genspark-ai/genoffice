# AI-modellek és beállítások

## Szolgáltatók és modellek

A modellek és kulcsok a Beállításokban adhatók meg (a Kezdőképernyő bal alsó sarkában lévő fióksoron):

![A Beállítások ablak](img/settings-general.png)

- **Genspark által üzemeltetett**: jelentkezz be (eszközkódos bejelentkezéssel), és már használhatod is — külön beállítás nélkül.
- **Egyéni végpontok (BYOK)**: a Beállítások ▸ AI oldalon protokollonként megadhatsz egy alap URL-t és egy API-kulcsot — OpenAI-kompatibilis, Anthropic, Gemini, DeepSeek, DashScope (qwen) és továbbiak. A kulcsok az alkalmazás beállítási fájljában tárolódnak ezen a gépen, és csak a kérések fejlécében kerülnek elküldésre. A beépített protokollok között az Atlas Cloud is szerepel.
- Képességenként más-más modellt választhatsz: csevegés/generálás, képgenerálás, képelemzés.
- **Kapcsolatteszt**: mentés előtt ellenőrizd a végpont elérhetőségét és a modellek láthatóságát.
- Az alap URL tartalmazhat elérési utat és lekérdezési paramétereket is (átjáró-stílusú címek); a végpontok útvonalait a program helyesen fűzi hozzá.

## CLI-integráció (Codex-osztályú)

- A Beállításokban megadhatod egy helyi CLI-program elérési útját (az ékezetes, nem ASCII karaktereket tartalmazó saját mappa és a ~ előtag is működik; a ~ automatikusan feloldódik); a Modellek észlelése lekérdezi a CLI elérhető modelljeit.
- Az ellenőrzés csak azt vizsgálja, hogy a fájl létezik-e — karakterkészlet-korlátozás nincs.

## Mikor lépnek életbe a módosítások

- A modell- és végpontmódosítások azonnal érvénybe lépnek; egy folyamatban lévő beszélgetés a következő körig még a régi beállítást használja.

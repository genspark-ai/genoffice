# Beállítások, nyelv, téma és MCP-integrációk

## A beállítások megnyitása

A Kezdőképernyő bal alsó sarkában lévő fióksor megnyitja a Beállítások panelt (kijelentkezve „Bejelentkezés” felirat áll rajta). Hat szakasza van: Fiók, AI-modell, AI média és keresés, Általános, Integrációk és Névjegy.

![Beállítások ▸ Általános: itt állíthatod be a nyelvet, a témát, az automatikus mentést és a használati statisztika kapcsolóját](img/settings-general.png)

Az AI-modell beállításairól külön cikk szól; az **AI média és keresés** szakaszban kapcsolhatod be szolgáltatónként a képgenerálást, a képelemzést, a videoelemzést, a webes keresést és a helyi fájlok keresését.

## Nyelv

- A Beállítások **22 felületi nyelvet** kínálnak: angol, egyszerűsített kínai, japán, koreai, francia, német, spanyol, thai, indonéz, orosz, arab, portugál, olasz, lengyel, cseh, holland, maláj, héber, hindi, hagyományos kínai, vietnámi, török.
- A váltás azonnal életbe lép, és megmarad; a natív menüsor is az új nyelven épül újra.

## Téma

Világos / sötét / rendszer szerinti. A „rendszer szerinti” az operációs rendszer megjelenését követi, és a szerkesztők villanás nélkül, vele együtt váltanak.

## Általános

- **Névtelen használati statisztikák küldése** — alapból bekapcsolva. Google Analytics 4-et használ, és elküldi a nyilvános IP-címedet és az átviteli metaadatokat; a dokumentumok tartalmát és a fájlneveket soha nem gyűjti, és minden esemény csak egy típust tartalmaz, például: „.docx megnyitása”. Itt bármikor kikapcsolhatod.
- **AI-oldalsáv pozíciója** (bal vagy jobb), **AI-panel szövegmérete** és **Helyesírás-ellenőrzés az AI-csevegésben**.
- **AI-panel megnyitása új dokumentumokban** — kikapcsolva az új dokumentum összecsukott panellel indul, és egy kattintással előhívható.
- Az **Összes dokumentum automatikus mentése** minden szerkesztőben alapértelmezetten bekapcsolja az automatikus mentést; egy-egy ablakban továbbra is kikapcsolhatod.
- **Mentési hely** egy **Módosítás** gombbal, valamint az **Alapértelmezett alkalmazás Office-dokumentumokhoz**, amellyel a .docx / .xlsx / .pptx fájlokat a GenOffice-hoz rendelheted.

## AI média és keresés

Ezek nem egyszerű kapcsolók — minden képességnél kiválasztod a kiszolgáló szolgáltatót, a szolgáltató kulcsát és alap URL-jét pedig csak egyszer kell megadnod, és közösen használatos:

- **Webes keresés**, **Képgenerálás**, **Képelemzés** és **Videoelemzés** — mindegyikhez szolgáltató, modell, kulcs és alap URL tartozik.
- A **Helyi fájlok keresése** ezen a gépen fut. Alatta a **Keresés újrarangsorolása Jevvel** alapértelmezés szerint **ki van kapcsolva**. Ha bekapcsolod, a 20 legjobb helyi találat részletei — dokumentumonként legfeljebb 1200 karakter, valamint a fájl- és mappanevek — az általad választott végpontra kerülnek (Jev OpenRouteren vagy TypeSafe-en keresztül, Perplexity, Liquid AI d1, Cloudflare, egy helyi motor vagy egyéni URL), ahol relevancia szerint újrarendezik őket. Kikapcsolva semmi nem hagyja el az eszközt.

## Névjegy

- **Verzió**, a projekt GitHub-hivatkozása és egy **Csillag a GitHubon** gomb.
- **Frissítési csatorna**: Stable vagy Beta. A módosítás azonnal életbe lép, és rögtön frissítést is keres; a Beta-telepítést nem állítja vissza Stable-re.

## Alapértelmezett alkalmazástársítások

A Beállításokban a GenOffice beállítható alapértelmezett kezelőként a .docx / .xlsx / .pptx / .pdf és hasonló fájltípusokhoz (ez rendszerszintű beállítás; ha a rendszer kéri, hagyd jóvá).

## Harmadik felek közleményei és frissítések

- Súgó ▸ Harmadik felek közleményei: az alkalmazással szállított nyílt forráskódú összetevők teljes licenclistája.
- Súgó ▸ Frissítések keresése: kézi ellenőrzést indít; ha van újabb verzió, felajánlja a telepítést.

## Bejelentkezés a Gensparkba

- A bejelentkezési pont (a Beállításokban vagy a felhőbeli projektlistában) **eszközkódos** folyamatot használ: a GenOffice megjelenít egy kódot, és megnyitja a bejelentkezési oldalt a böngészőben; ha végeztél, a folyamat automatikusan folytatódik.
- A bejelentkezésre kizárólag ezekhez van szükség: a felhőbeli projektlistához és a Genspark által üzemeltetett modellekhez. Enélkül is működik minden helyi funkció és az egyéni modellek.
- A kijelentkezés egyetlen kattintás a Beállításokban.

## MCP-integráció (haladó felhasználóknak / AI-klienseknek)

Az **Integrációk** az a panel, amely összeköti a GenOffice-ot egy kódoló ügynökkel, és külön cikk is szól róla: Kódoló ügynök csatlakoztatása. Röviden: válassz egy módot (parancssor vagy MCP), kövesd az adott szakaszt, majd indíts új csevegést, és kérdezz.

![Beállítások ▸ Integrációk: a három lépés, majd a képességek sorai és az MCP-beállítások](img/settings-integrations.png)

A **Helyi HTTP-kiszolgáló** alatt az alkalmazás maga is futtathatja a kiszolgálót — van hozzá egy engedélyezési kapcsoló és egy port —, a **Speciális** rész pedig az állapotellenőrző URL-t és a naplófájlt is megjeleníti. Csak a localhoston figyel.

## Parancssori gyorsreferencia

| Parancs            | Mit csinál                        |
| ------------------ | --------------------------------- |
| `genoffice <file>` | fájl megnyitása                   |
| `genoffice mcp`    | a helyi MCP-kiszolgáló elindítása |
| `genoffice --help` | az összes parancs és kapcsoló     |

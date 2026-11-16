# A GenOffice telepítése

Minden build elérhető a [Releases oldalon](https://github.com/genspark-ai/genoffice/releases/latest). Mindegyik a `main` ágból készül; a macOS- és Windows-telepítők alá vannak írva.

## Válaszd ki a gépednek megfelelő fájlt

| Platform                             | Követelmények                                         | Fájl                  |
| ------------------------------------ | ----------------------------------------------------- | --------------------- |
| **macOS** — Apple Silicon            | macOS 11+                                             | `.dmg` (arm64)        |
| **macOS** — Intel                    | macOS 11+                                             | `.dmg` (x64)          |
| **Windows** — a legtöbb PC           | Windows 10+, Intel/AMD                                | `-x64.exe` telepítő   |
| **Windows** on Arm                   | Windows 11 on Arm (Snapdragon X és hasonlók)          | `-arm64.exe` telepítő |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 vagy újabb)         | `.deb`                |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm`                |
| **Linux** — minden más               | x86_64, glibc 2.34+, FUSE 2                           | `.AppImage`           |

A régebbi verziók is megmaradnak a Releases oldalon.

## macOS-en

Nyisd meg a `.dmg` fájlt, és húzd a **GenOffice**-ot az Alkalmazások mappába. Az Intel-es buildnél az első indításkor meg kell erősítened a megnyitást: kattints jobb gombbal az alkalmazásra az Alkalmazások mappában ▸ Megnyitás. Ez csak annyit jelent, hogy a Gatekeeper aláíratlannak tűnő lemezképpel találkozott, nem azt, hogy a letöltés megsérült.

## Windowson

Futtasd az `.exe` telepítőt. Felveszi a GenOffice-ot a Start menübe, és társítja a `.docx`, `.xlsx`, `.pptx`, `.pdf` és hasonló fájltípusokat, így egy fájlra duplán kattintva megnyílik a GenOffice-ban.

## Linuxon

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — ott fut, ahol van, telepítési lépés nélkül:

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

Ehhez a FUSE 2 futtatókörnyezet szükséges. Ha nem szeretnéd telepíteni, indítsd a `--appimage-extract-and-run` kapcsolóval, és a sandbox a kicsomagolt mappából működik.

## A használt build

A **Beállítások ▸ Névjegy** panelen látod, melyik buildet használod, melyik csatornát követi az alkalmazás, és hol található a projekt. Ha ezen a csatornán újabb build jelenik meg, ugyanez a panel felajánlja a frissítést; Flatpak-, Nix- vagy Docker-telepítésnél ez nincs így, mert ezeket az az eszköz cseréli le, amellyel telepítetted őket.

![Beállítások ▸ Névjegy, amelyen látható a telepített verzió, az alkalmazás által követett frissítési csatorna és a projekt GitHub-hivatkozása](img/install.png)

## Három további telepítési mód

Mindhárom a közzétett kiadási fájlokat csomagolja be ahelyett, hogy újraépítené az alkalmazást, így a hivatalos binárisokat követi.

- **Flatpak** — a `.deb` csomagot a Flatpak `extra-data` mechanizmusán keresztül telepíti, és a Zypak alatt futtatja, amely az Electron saját sandboxát igazítja a Flatpakhoz. Az `/app` könyvtár csak olvasható, ezért a frissítés a Flatpakon keresztül történik, nem az alkalmazás beépített frissítési ablakával.
- **Nix** — egy csomagkifejezés, amelyet a forráskönyvtárból építhetsz a `nix-build packaging/nix` paranccsal; részletek: `packaging/nix`.
- **Docker** — grafikus felület nélküli image kötegelt konvertáláshoz (`packaging/docker`). Office-, Markdown- és HTML-dokumentumok teljes könyvtárfáját alakítja PDF-fé az alkalmazás saját exportálási folyamatán keresztül, így a kimenet ugyanaz, mint amit a Fájl ▸ Exportálás adna. Ez egy CLI-segédprogram, nem szerver.

## A parancssor is jár hozzá

Minden telepítés tartalmaz egy `genoffice` parancsot, amely ugyanazokat a motorokat használja, mint az ablak, így a kettő soha nem kezeli eltérően ugyanazt a fájlt. macOS-en és Windowson ez az alkalmazáscsomagon belül található; a `genoffice install-cli` hozzáadja a `PATH`-hoz. A **Parancssor és ügynökök** című súgótémából megtudhatod, mire képes.

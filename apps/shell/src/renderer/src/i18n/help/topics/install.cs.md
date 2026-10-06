# Instalace GenOffice

Každá sestavení je zveřejněno na [stránce Releases](https://github.com/genspark-ai/genoffice/releases/latest). Všechna pocházejí z `main`; instalátory pro macOS a Windows jsou podepsané.

## Vyberte soubor pro svůj počítač

| Platforma | Požadavky | Soubor |
| --- | --- | --- |
| **macOS** — Apple Silicon | macOS 11+ | `.dmg` (arm64) |
| **macOS** — Intel | macOS 11+ | `.dmg` (x64) |
| **Windows** — většina počítačů | Windows 10+, Intel/AMD | instalátor `-x64.exe` |
| **Windows** s architekturou Arm | Windows 11 on Arm (Snapdragon X a podobné) | instalátor `-arm64.exe` |
| **Linux** — Debian / Ubuntu | x86_64, glibc 2.34+ (Ubuntu 22.04 nebo novější) | `.deb` |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm` |
| **Linux** — cokoliv jiného | x86_64, glibc 2.34+, FUSE 2 | `.AppImage` |

Starší verze zůstávají na stránce Releases.

## V systému macOS

Otevřete soubor `.dmg` a přetáhněte **GenOffice** do složky Aplikace. U sestavení Intel vás při prvním spuštění požádá o potvrzení otevření: klikněte na aplikaci v Aplikacích pravým tlačítkem ▸ Otevřít. To je Gatekeeper, který narazí na obraz disku, který vypadá jako nepodepsaný, a ne poškozené stažení.

## V systému Windows

Spusťte instalátor `.exe`. Umístí GenOffice do nabídky Start a zaregistruje `.docx`, `.xlsx`, `.pptx`, `.pdf` a podobné, takže dvojklik na soubor jej otevře.

## V systému Linux

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — běží tam, kde je, bez žádného instalačního kroku:

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

Potřebuje běhové prostředí FUSE 2. Pokud jej nechcete instalovat, spusťte jej s `--appimage-extract-and-run` a izolovaný proces poběží z rozbalené složky.

## Tři další cesty

Všechny tři obalí zveřejněné artefakty místo toho, aby aplikaci znovu sestavovaly, takže sledují oficiální binární soubory.

- **Flatpak** — nainstaluje `.deb` pomocí mechanismu `extra-data` ve Flatpaku a spustí jej pod Zypakem, který přizpůsobuje vlastní izolaci Electronu. Strom `/app` je jen pro čtení, takže aktualizace jdou přes Flatpak, ne přes dialog aktualizace v aplikaci.
- **Nix** — výraz balíčku, který sestavíte z checkoutu příkazem `nix-build packaging/nix`; viz `packaging/nix`.
- **Docker** — obraz pro dávkovou převod bez rozhraní (`packaging/docker`). Převede celý strom dokumentů Office, Markdown a HTML do PDF přes vlastní exportní kanál aplikace, takže výstup je stejný, jaký by dalo Soubor ▸ Exportovat. Je to nástroj CLI, ne server.

## Příkazová řádka je součástí

Každá instalace obsahuje příkaz `genoffice`, který mluví se stejnými enginy jako okno, takže se nikdy neshodnou v tom, jaký soubor je. V systémech macOS a Windows je uvnitř balíčku aplikace; `genoffice install-cli` jej umístí do vaší `PATH`. Co umí, se dozvíte v kapitole **Příkazová řádka a agenti**.

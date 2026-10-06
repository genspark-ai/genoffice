# GenOffice installeren

Elke build staat op de [Releases-pagina](https://github.com/genspark-ai/genoffice/releases/latest). Ze komen allemaal uit `main`; de installateurs voor macOS en Windows zijn ondertekend.

## Kies het bestand voor je computer

| Platform | Vereisten | Bestand |
| --- | --- | --- |
| **macOS** — Apple Silicon | macOS 11+ | `.dmg` (arm64) |
| **macOS** — Intel | macOS 11+ | `.dmg` (x64) |
| **Windows** — de meeste pc's | Windows 10+, Intel/AMD | installateur `-x64.exe` |
| **Windows** op Arm | Windows 11 on Arm (Snapdragon X en vergelijkbaar) | installateur `-arm64.exe` |
| **Linux** — Debian / Ubuntu | x86_64, glibc 2.34+ (Ubuntu 22.04 of nieuwer) | `.deb` |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm` |
| **Linux** — al het andere | x86_64, glibc 2.34+, FUSE 2 | `.AppImage` |

Oudere versies blijven op de Releases-pagina staan.

## Op macOS

Open het `.dmg` en sleep **GenOffice** naar Programma's. Bij een Intel-build vraagt de eerste keer starten om bevestiging: klik met de rechtermuis op de app in Programma's ▸ Open. Dat is Gatekeeper die een schijfkopie tegenkomt die niet-ondertekend lijkt, geen beschadigde download.

## Op Windows

Voer de installateur `.exe` uit. Die zet GenOffice in het Startmenu en registreert `.docx`, `.xlsx`, `.pptx`, `.pdf` en soortgelijke, zodat een dubbelklik op een bestand het opent.

## Op Linux

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — draait waar het staat, zonder installatiestap:

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

Het heeft de FUSE 2-runtime nodig. Wil je die niet installeren, start het dan met `--appimage-extract-and-run`; de sandbox werkt vanuit de uitgepakte map.

## Nog drie routes

Ze pakken alle drie de gepubliceerde artefacten in plaats van de app opnieuw te bouwen, dus ze volgen de officiële binaire bestanden.

- **Flatpak** — installeert de `.deb` via het `extra-data`-mechanisme van Flatpak en draait die onder Zypak, dat de eigen sandbox van Electron aanpast. De `/app`-structuur is alleen-lezen, dus updates lopen via Flatpak en niet via het updatevenster in de app.
- **Nix** — een pakketuitdrukking die je vanuit een checkout bouwt met `nix-build packaging/nix`; zie `packaging/nix`.
- **Docker** — een headless image voor batchconversie (`packaging/docker`). Het zet een hele boom aan Office-, Markdown- en HTML-documenten om naar PDF via de eigen exportpijplijn van de app, dus de uitvoer is wat Bestand ▸ Exporteren zou geven. Het is een CLI-hulpmiddel, geen server.

## De opdrachtregel zit erbij

Elke installatie bevat een `genoffice`-commando dat met dezelfde engines praat als het venster, zodat de twee nooit over een bestand het oneens zijn. Op macOS en Windows zit het in de app-bundel; `genoffice install-cli` zet het in je `PATH`. Wat het kan doen, lees je in **Opdrachtregel en agents**.

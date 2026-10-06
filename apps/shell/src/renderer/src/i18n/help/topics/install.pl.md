# Instalacja GenOffice

Każda kompilacja jest publikowana na [stronie Releases](https://github.com/genspark-ai/genoffice/releases/latest). Wszystkie pochodzą z `main`; instalatory dla macOS i Windows są podpisane.

## Wybierz plik dla swojego komputera

| Platforma | Wymagania | Plik |
| --- | --- | --- |
| **macOS** — Apple Silicon | macOS 11+ | `.dmg` (arm64) |
| **macOS** — Intel | macOS 11+ | `.dmg` (x64) |
| **Windows** — większość komputerów | Windows 10+, Intel/AMD | instalator `-x64.exe` |
| **Windows** na Arm | Windows 11 on Arm (Snapdragon X i podobne) | instalator `-arm64.exe` |
| **Linux** — Debian / Ubuntu | x86_64, glibc 2.34+ (Ubuntu 22.04 lub nowsze) | `.deb` |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm` |
| **Linux** — wszystko inne | x86_64, glibc 2.34+, FUSE 2 | `.AppImage` |

Starsze wersje pozostają na stronie Releases.

## W systemie macOS

Otwórz plik `.dmg` i przeciągnij **GenOffice** do katalogu Aplikacje. W kompilacji Intel przy pierwszym uruchomieniu trzeba potwierdzić otwarcie: kliknij prawym przyciskiem aplikację w Aplikacjach ▸ Otwórz. To Gatekeeper reaguje na obraz dysku wyglądający na niepodpisany, a nie uszkodzone pobranie.

## W systemie Windows

Uruchom instalator `.exe`. Umieszcza GenOffice w menu Start i rejestruje `.docx`, `.xlsx`, `.pptx`, `.pdf` i podobne, więc dwuklik na plik go otwiera.

## W systemie Linux

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — działa tam, gdzie leży, bez żadnego etapu instalacji:

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

Wymaga środowiska uruchomieniowego FUSE 2. Jeśli wolisz go nie instalować, uruchom go z `--appimage-extract-and-run`, a sandbox zadziała z katalogu wypakowanego.

## Trzy inne drogi

Wszystkie trzy opakowują opublikowane artefakty zamiast przebudowywać aplikację, więc pozostają zgodne z oficjalnymi plikami binarnymi.

- **Flatpak** — instaluje `.deb` przez mechanizm `extra-data` Flatpaka i uruchamia go pod Zypakiem, który dostosowuje własny sandbox Electrona. Drzewo `/app` jest tylko do odczytu, więc aktualizacje przechodzą przez Flatpak, a nie przez okno aktualizacji w aplikacji.
- **Nix** — wyrażenie pakietu, które budujesz z checkoutu poleceniem `nix-build packaging/nix`; zobacz `packaging/nix`.
- **Docker** — obraz headless do konwersji wsadowej (`packaging/docker`). Konwertuje całe drzewo dokumentów Office, Markdown i HTML do PDF przez własny potok eksportu aplikacji, więc wynik jest taki sam jak po Plik ▸ Eksportuj. To narzędzie CLI, a nie serwer.

## Wiersz poleceń jest w zestawie

Każda instalacja zawiera polecenie `genoffice`, które komunikuje się z tymi samymi silnikami co okno, więc nigdy nie rozbieżą się w ocenie pliku. W systemach macOS i Windows znajduje się w pakiecie aplikacji; `genoffice install-cli` umieszcza je w `PATH`. Co potrafi, dowiesz się w rozdziale **Wiersz poleceń i agenci**.

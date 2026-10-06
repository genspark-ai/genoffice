# Installare GenOffice

Ogni build viene pubblicato sulla [pagina Releases](https://github.com/genspark-ai/genoffice/releases/latest). Provengono tutti da `main`; i programmi di installazione per macOS e Windows sono firmati.

## Scegli il file giusto per il tuo computer

| Piattaforma | Requisiti | File |
| --- | --- | --- |
| **macOS** — Apple Silicon | macOS 11+ | `.dmg` (arm64) |
| **macOS** — Intel | macOS 11+ | `.dmg` (x64) |
| **Windows** — la maggior parte dei PC | Windows 10+, Intel/AMD | programma di installazione `-x64.exe` |
| **Windows** su Arm | Windows 11 on Arm (Snapdragon X e simili) | programma di installazione `-arm64.exe` |
| **Linux** — Debian / Ubuntu | x86_64, glibc 2.34+ (Ubuntu 22.04 o successivo) | `.deb` |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm` |
| **Linux** — qualsiasi altra cosa | x86_64, glibc 2.34+, FUSE 2 | `.AppImage` |

Le versioni più vecchie restano sulla pagina Releases.

## Su macOS

Apri il file `.dmg` e trascina **GenOffice** in Applicazioni. Su una build Intel il primo avvio ti chiede di confermare l'apertura: fai clic con il tasto destro sull'app in Applicazioni ▸ Apri. Non è un download danneggiato: è Gatekeeper che incontra un'immagine disco che sembra non firmata.

## Su Windows

Esegui il programma di installazione `.exe`. Inserisce GenOffice nel menu Start e registra `.docx`, `.xlsx`, `.pptx`, `.pdf` e simili, così un doppio clic su un file lo apre.

## Su Linux

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — funziona da dove si trova, senza alcuna fase di installazione:

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

Richiede il runtime FUSE 2. Se preferisci non installarlo, avvialo con `--appimage-extract-and-run`: la sandbox funziona dalla directory estratta.

## Altre tre vie

Tutte e tre avvolgono gli artefatti pubblicati invece di ricompilare l'app, così restano allineate ai binari ufficiali.

- **Flatpak** — installa il `.deb` tramite il meccanismo `extra-data` di Flatpak e lo esegue sotto Zypak, che adatta la sandbox propria di Electron. L'albero `/app` è in sola lettura, quindi gli aggiornamenti passano da Flatpak e non dalla finestra di aggiornamento dell'app.
- **Nix** — un'espressione di pacchetto che costruisci da un checkout con `nix-build packaging/nix`; vedi `packaging/nix`.
- **Docker** — un'immagine headless per la conversione in blocco (`packaging/docker`). Converte un intero albero di documenti Office, Markdown e HTML in PDF attraverso la pipeline di esportazione dell'app, quindi l'output è quello che produrrebbe File ▸ Esporta. È un'utilità CLI, non un server.

## La riga di comando è inclusa

Ogni installazione include il comando `genoffice`, che comunica con gli stessi motori della finestra, così i due non si contraddicono mai su un file. Su macOS e Windows si trova dentro il bundle dell'app; `genoffice install-cli` lo mette sul tuo `PATH`. Vedi **Riga di comando e agenti** per scoprire cosa può fare.

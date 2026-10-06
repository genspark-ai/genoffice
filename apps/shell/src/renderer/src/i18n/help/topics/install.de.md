# GenOffice installieren

Jeder Build wird auf der [Releases-Seite](https://github.com/genspark-ai/genoffice/releases/latest) veröffentlicht. Alle stammen aus `main`; die Installer für macOS und Windows sind signiert.

## Die Datei für Ihren Rechner wählen

| Plattform | Voraussetzungen | Datei |
| --- | --- | --- |
| **macOS** — Apple Silicon | macOS 11+ | `.dmg` (arm64) |
| **macOS** — Intel | macOS 11+ | `.dmg` (x64) |
| **Windows** — die meisten PCs | Windows 10+, Intel/AMD | `-x64.exe` Installationsprogramm |
| **Windows** auf Arm | Windows 11 auf Arm (Snapdragon X und Ähnliches) | `-arm64.exe` Installationsprogramm |
| **Linux** — Debian / Ubuntu | x86_64, glibc 2.34+ (Ubuntu 22.04 oder neuer) | `.deb` |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm` |
| **Linux** — alles andere | x86_64, glibc 2.34+, FUSE 2 | `.AppImage` |

Ältere Versionen bleiben auf der Releases-Seite.

## Unter macOS

Öffnen Sie die `.dmg` und ziehen Sie **GenOffice** in Applications. Bei einem Intel-Build fragt der erste Start das Öffnen ab: Rechtsklick auf die App in Applications ▸ Öffnen. Das ist Gatekeeper bei einem scheinbar nicht signierten Datenträgerabbild und keine beschädigte Datei.

## Unter Windows

Führen Sie das Installationsprogramm `.exe` aus. Es legt GenOffice im Startmenü ab und registriert `.docx`, `.xlsx`, `.pptx`, `.pdf` und Verwandte, sodass ein Doppelklick auf eine Datei sie öffnet.

## Unter Linux

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — läuft ohne Installationsschritt direkt von dort:

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

Er benötigt die FUSE-2-Laufzeit. Wenn Sie sie lieber nicht installieren, starten Sie ihn mit `--appimage-extract-and-run`; die Sandbox arbeitet dann aus dem entpackten Verzeichnis.

## Drei weitere Wege

Alle drei verpacken die veröffentlichten Artefakte, statt die App neu zu bauen; sie folgen damit den offiziellen Binärdateien.

- **Flatpak** — installiert die `.deb` über den `extra-data`-Mechanismus von Flatpak und führt sie unter Zypak aus, das die eigene Sandbox von Electron anpasst. Der Baum `/app` ist schreibgeschützt, daher laufen Updates über Flatpak und nicht über den Aktualisierungsdialog in der App.
- **Nix** — ein Paketausdruck, den Sie aus einem Checkout mit `nix-build packaging/nix` bauen; siehe `packaging/nix`.
- **Docker** — ein Headless-Image für die Stapelkonvertierung (`packaging/docker`). Es konvertiert einen ganzen Baum aus Office-, Markdown- und HTML-Dokumenten über die eigene Export-Pipeline der App nach PDF, sodass das Ergebnis dem von Datei ▸ Exportieren entspricht. Es ist ein CLI-Werkzeug, kein Server.

## Die Kommandozeile ist dabei

Jede Installation bringt den Befehl `genoffice` mit, der mit denselben Engines spricht wie das Fenster: Die beiden können sich bei einer Datei also nicht widersprechen. Unter macOS und Windows liegt er im App-Bundle; `genoffice install-cli` legt ihn in Ihren `PATH`. Was er kann, steht in **Kommandozeile und Agenten**.

# Installer GenOffice

Chaque build est publié sur la [page Releases](https://github.com/genspark-ai/genoffice/releases/latest). Tous proviennent de `main` ; les installeurs macOS et Windows sont signés.

## Choisir le fichier pour votre machine

| Plateforme | Prérequis | Fichier |
| --- | --- | --- |
| **macOS** — Apple Silicon | macOS 11+ | `.dmg` (arm64) |
| **macOS** — Intel | macOS 11+ | `.dmg` (x64) |
| **Windows** — la plupart des PC | Windows 10+, Intel/AMD | installateur `-x64.exe` |
| **Windows** sur Arm | Windows 11 sur Arm (Snapdragon X et similaires) | installateur `-arm64.exe` |
| **Linux** — Debian / Ubuntu | x86_64, glibc 2.34+ (Ubuntu 22.04 ou plus récent) | `.deb` |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm` |
| **Linux** — tout le reste | x86_64, glibc 2.34+, FUSE 2 | `.AppImage` |

Les versions antérieures restent sur la page Releases.

## Sous macOS

Ouvrez le `.dmg` et faites glisser **GenOffice** dans Applications. Sur une version Intel, le premier lancement demande de confirmer l’ouverture : clic droit sur l’application dans Applications ▸ Ouvrir. C’est Gatekeeper qui rencontre une image disque sans signature, pas un téléchargement endommagé.

## Sous Windows

Lancez l’installeur `.exe`. Il place GenOffice dans le menu Démarrer et enregistre `.docx`, `.xlsx`, `.pptx`, `.pdf` et formats voisins, si bien qu’un double-clic sur un fichier l’ouvre.

## Sous Linux

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — s’exécute sur place, sans étape d’installation :

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

Il lui faut l’environnement d’exécution FUSE 2. Si vous préférez ne pas l’installer, lancez-le avec `--appimage-extract-and-run` : le bac à sable fonctionne alors depuis le répertoire extrait.

## Trois autres voies

Les trois encapsulent les artefacts publiés au lieu de reconstruire l’application : elles suivent donc les binaires officiels.

- **Flatpak** — installe le `.deb` via le mécanisme `extra-data` de Flatpak et l’exécute sous Zypak, qui adapte le bac à sable d’Electron. L’arborescence `/app` est en lecture seule : les mises à jour passent donc par Flatpak et non par la boîte de dialogue de mise à jour de l’application.
- **Nix** — une expression de paquet que vous compilez depuis une copie de dépôt avec `nix-build packaging/nix` ; voir `packaging/nix`.
- **Docker** — une image de conversion par lots sans interface (`packaging/docker`). Elle convertit un arbre entier de documents Office, Markdown et HTML en PDF via la chaîne d’export de l’application : le résultat est donc celui de Fichier ▸ Exporter. C’est un utilitaire CLI, pas un serveur.

## La ligne de commande est incluse

Chaque installation livre une commande `genoffice` qui parle aux mêmes moteurs que la fenêtre : les deux ne peuvent donc pas se contredire sur un fichier. Sous macOS et Windows, elle se trouve dans le bundle de l’application ; `genoffice install-cli` la met sur votre `PATH`. Pour ce qu’elle sait faire, voir **Ligne de commande et agents**.

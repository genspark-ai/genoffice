# GenOffice on ChromeOS

ChromeOS runs Linux applications in an integrated Debian VM ("Crostini", the
Terminal app's Linux development environment). GenOffice ships as a standard
Linux desktop package, so it runs there like any other Linux app — no separate
ChromeOS build exists or is needed. This page covers what works, how to set it
up, and where the limits are.

## The Linux desktop app (Crostini)

The full application: every editor, offline by default, files stay local.

Requirements:

- A Chromebook with Linux support (most 2019+ devices; check
  `Settings → About ChromeOS` for the "Developers → Linux development
  environment" toggle). The published Linux packages are x86_64, so use an
  Intel or AMD Chromebook.
- Roughly 6 GB of free space for the container plus ~500 MB for the app.

Install (Debian/Ubuntu-style):

```sh
sudo apt install ./genoffice_<version>_amd64.deb
genoffice  # or launch from the app drawer
```

The deb pulls in the runtime libraries it needs (`libgtk-3-0`, `libnss3`,
`libsecret-1-0`, and friends — all present in the Debian 12 image Crostini
ships). The self-contained AppImage is an alternative that installs nothing:

```sh
chmod +x GenOffice-<version>.AppImage && ./GenOffice-<version>.AppImage
```

Files: Crostini shares the device's "My files" folder at `~/` inside Linux, so
documents opened from or saved to `~/MyFiles/` are visible to the ChromeOS
Files app and to Android apps. Saving into other Linux paths stays inside the
container.

Printing: Crostini forwards CUPS queues from the host; the print dialog lists
the ChromeOS-configured printers.

Backups: the Linux container's files are included in ChromeOS device backups
only partially — keep important documents in `~/MyFiles/` so they live on the
host filesystem.

Display: the container's Sommelier bridges Linux windows to the ChromeOS
desktop, over native Wayland or X11; GenOffice needs no display configuration.

## What does not work / what does not exist

- **Play Store / Android app**: GenOffice is not an Android application; ARC
  (the Android runtime) is not a supported install path. The Electron desktop
  architecture does not translate to Android, and this is not planned.
- **Google Drive sync**: no Drive integration is built in. Documents live in
  the Linux filesystem; ChromeOS can surface Drive files through the Files app,
  but the Linux container does not mount Drive. Export/save manually, or keep
  working files in `MyFiles`.

## Privacy notes

Source and fork builds run with analytics disabled and without the
update/font-catalog endpoints; the official release injects anonymous usage
analytics (GA4 Measurement Protocol) and an optional font catalog — the same
behavior as the Windows/macOS/Linux desktop builds, nothing ChromeOS-specific
is collected. No document content leaves the machine except through the
features the user explicitly invokes (AI providers, updates, font downloads).

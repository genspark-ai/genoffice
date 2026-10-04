# Fonts: system faces and downloadable families

## The font lists

The font controls live on each editor's Home tab (below: the font group in Slides):

![The font group in the Slides ribbon](img/fonts.en.png)

Each editor's font picker merges: locally installed fonts + platform-common candidates (the usual Windows and macOS families, Western and CJK, including the localized names CJK fonts report on their home systems). When the system can enumerate local fonts the list groups by what actually exists; otherwise it degrades to the full candidate list.

## Downloadable fonts (Slides)

- The Slides font picker opens a **font catalog**: a curated OFL selection spanning many scripts; every family ships regular and bold weights.
- Picking one downloads and installs it from the CDN (checksum-pinned) into the app's font store — available to every document afterwards, no system-level install.
- Installs live in an app-private directory and go away with the app.

## CJK typography

- CJK fonts in Docs get Word-aligned punctuation compression and kinsoku line rules (see the Docs chapter).
- The catalog's CJK families cover Simplified/Traditional Chinese, Japanese and Korean, serif and sans.

## Managing local fonts

- The install-local-font-files entry loads .ttf/.otf files from disk into the app font store.
- The store archives by family; multiple weights of one family coexist.

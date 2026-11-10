# Shared bundled fonts

Font files that more than one app ships. Each app used to carry its own copy;
keeping a single set here means one download per package, one licence record,
and one place for the build-time patch to write to.

| Font       | Files           | License | Used by              |
| ---------- | --------------- | ------- | -------------------- |
| Carlito GO | `Carlito-*.ttf` | OFL 1.1 | docs, sheets, slides |

"Carlito GO" is Carlito 1.103 (LibreOffice's metric-compatible Calibri
substitute) with Vietnamese stacked diacritics rebuilt by
`tools/patch-carlito-vi.py`, renamed per OFL 1.1 §2 ("Carlito" is a Reserved
Font Name). The full OFL text lives in `apps/docs/src/renderer/fonts/LICENSE-OFL.txt`
and the per-font notes in `apps/docs/src/renderer/fonts/README.md`.

Reference the files through the package export, from CSS or from a main-process
asset import:

```css
src: url('@genoffice/ui/fonts/Carlito-Regular.ttf');
```

```ts
import carlitoRegular from '@genoffice/ui/fonts/Carlito-Regular.ttf?asset'
```

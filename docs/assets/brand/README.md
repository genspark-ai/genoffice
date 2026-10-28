# GenOffice brand assets

Master: `logo.svg` (the same file lives at `apps/shell/build/logo.svg` — edit
that one and regenerate; this copy is the reference for non-app uses).

## Mark

"Four pages, one spark": four folded pages — one per editor, each in its app
accent (docs `#0f7fff`, sheets `#16a34a`, slides `#f97316`, pdf `#ef4444`) —
pinwheel around a four-point AI spark in negative space.

## Rules

- Clear space: keep at least 12% of the icon's width empty on every side.
- The mark ships on transparent; never put it back on a black tile (the
  retired hummingbird tile did that and vanished on dark taskbars).
- Minimum sizes: the full-color mark stays legible to 16px; below that use
  the 16px asset as-is instead of downscaling further.
- File-type icons (`apps/shell/src/renderer/src/assets/file-*.svg`) keep the
  same accent palette; regenerate OS association icons with
  `node tools/gen-file-association-icons.mjs` after editing them.

## Regenerate the app icon set

```
node_modules/.bin/electron tools/build-icons.cjs
```

Outputs: `apps/shell/build/icon.png`, `icon-mac.png`, `icon.ico`, `icon.icns`,
`icons/<size>.png`, and the renderer copy `src/assets/app-icon.png`
(copied from `icon.png`).

# Optimization journal

Five self-driven optimization rounds on the desktop feature set. One direction
per round, never repeated. Each entry records the direction, what the
measurements said, what changed, and how it was verified.

## Round 1 — Performance (wallpaper + desktop rendering)

Direction: wallpaper PNG payloads and decode cost. The six built-ins are
1920×1080 PNGs generated procedurally; check their byte size, the data-URL IPC
transfer shape, and the desktop's render path for avoidable work.

Findings: PNGs are 59–297 KB each (fine); `wallpaperImage` is demand-loaded
(active backdrop + center-open catalog only, with pending-set dedup — already
good). The real hot spot was the starfield wallpaper's canvas: a 60fps
requestAnimationFrame loop redrawing up to 160 stars whose twinkle sine runs
at t·0.001 — far slower than the frame rate.

Change: the twinkle now draws every 5th frame (~12fps — visually identical,
~80% less draw work). Verified: 89/89 desktop tests pass; the rest of the
animation budget (aurora blobs, icon fade-in) was already transform/opacity
only.

## Round 2 — Visual polish (label legibility)

Direction: screenshot review over a dark wallpaper. Found: with a dark
built-in wallpaper and labelColor '' (auto), icon labels stayed in the theme's
dark ink — nearly invisible on the dark art.

Change: automatic ink now follows the wallpaper — builtins carry a known
darkness set, custom wallpapers get their average luminance sampled once
through a 32×32 canvas, and the CSS aurora follows the active theme. Light ink
plus the opposite scrim lands over dark art; light art keeps the theme pair.
Verified: two new tests (dark builtin flips, light builtin doesn't) + a
starfield screenshot with every label readable.

## Round 3 — Interaction (OS-desktop keyboard muscle memory)

Direction: what a real desktop does with a keyboard. Missing: arrow-key
movement between icons and type-ahead selection.

Change: ArrowUp/Down step through the flow order, ArrowLeft/Right jump a whole
visible column (reads the measured grid), typing a printable run within 800ms
selects the first icon whose label starts with it. The desktop only listens
when it owns the keyboard (surface or an icon focused — sidebar inputs and the
search box are untouched), dialogs and open menus take precedence.

Verified: three new tests (arrow movement, type-ahead jump, input immunity);
56/56 desktop-view tests.

## Round 4 — Accessibility (the icon grid is a widget, not a pile of divs)

Direction: what a screen reader hears on the desktop.

Change: the icon list is now a real listbox — `role="listbox"` with
`role="option"` children carrying `aria-selected`, `role="presentation"` on
the positioning wrappers. Selection state is announced, arrow/type-ahead
navigation (round 3) drives it. Verified: new semantics test + 57/57
desktop-view tests.

## Round 5 — Code health (one file, one job)

Direction: DesktopView.tsx had grown to ~2200 lines mixing the surface, the
wallpaper-center dialog, and the properties dialogs.

Change: extracted `DesktopWallpaperCenter.tsx` (345 lines: tiles, fit row,
label-color row, carousel, HEX draft state) and `DesktopProperties.tsx`
(281 lines: both properties dialogs with their data fetches). DesktopView.tsx
(1734 lines) keeps the surface: grid, drag, menus, backdrop, keyboard, dialog
open state. DOM output is byte-identical by construction (class names, roles,
aria all unchanged) — proven by the untouched test suite staying green:
103/103 desktop tests, 1441 shell tests.

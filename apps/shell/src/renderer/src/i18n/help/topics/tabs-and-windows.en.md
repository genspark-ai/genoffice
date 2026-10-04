# Tabs and window management

Every open file shares one window; the tab strip at the top switches between them, browser-style.

![The tab strip: Home plus two documents](img/tabs.en.png)

## Basics

- **Switch**: click a tab, or scroll the wheel over the strip to cycle.
- **Close**: the × on the tab, or ⌘W/ctrl+W. Unsaved changes ask first; the all-tabs menu can close every other / right-hand tab at once.
- **New**: the + at the right end of the strip.
- **Activation is immediate**: pressing a tab switches as soon as it is pressed, without waiting for the click to finish.

## Rename: double-click a tab

**Double-click any file tab** and the title turns into an inline input: Enter commits, Escape cancels, blur commits; the Enter that confirms an IME candidate is not mistaken for a commit. The rename goes through the same safety checks as Home's row rename (illegal characters, name conflicts), renames the file on disk and refreshes the recents list.

## Drag to reorder

Hold and drag a tab sideways to reorder; neighbours slide aside live and the order settles on release. A 4-pixel dead zone keeps plain clicks from wiggling tabs.

## Context menu

Right-click a tab for: new tab, close, close others, close to the right, duplicate the file and more (localized with the app language).

## Tear off into a window

**Drag a tab out of the strip** (or use Open in new window) and it becomes its own window carrying the live document; drag it back to dock it into the strip again. A detached window is a full peer — keep editing and saving there.

## The all-tabs list

When tabs overflow, the ▾ at the right end of the strip opens the full list (a native menu, never covered by the content area); pick with the arrow keys.

## The toolbar row inside every editor

Each editor tab has a toolbar row at the top (exact layout varies a little per editor):

- **Save** (⌘S/ctrl+S) and **Save As**.
- **Undo / Redo**: multi-level; each editor keeps its own history.
- **Find** (ctrl+F): opens that editor's find panel.
- **AutoSave** toggle: when on, changes are written to disk on a timer; when off, only manual saves write (the unsaved marker on the tab/title reminds you).

## The Home tab

The leftmost Home tab cannot be closed; to get back from any editor, click it or use File ▸ Home.

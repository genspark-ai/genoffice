# The HTML editor

The HTML editor opens .html / .htm with two modes: **preview** (the rendered page) and **source**.

- **Preview**: true rendering; relative stylesheets and images load next to the file.
- **Preview inspector**: click to select an element, double-click to edit its text in place, toolbar delete, and Ask AI against the selection.
- **Source mode**: edit the HTML; ctrl+F find, Replace All saves the rewritten markup.
- **Saving**: byte-faithful (BOM/CRLF/trailing newline preserved); unmodified saves do not rewrite.
- **Zoom**: ctrl+wheel / pinch scales the preview; ctrl+Z inside the preview undoes the last edit.

## The toolbar

Click any element in the preview and a toolbar floats above it:

![The floating toolbar over a selected element](img/html-toolbar.png)

- **File & history**: Save, Save As, Undo, Redo, Find; the **AutoSave** toggle writes changes on a timer.
- **Preview / Source** switch; **Present** shows the page full-screen.
- **Formatting**: bold, italic, font size up/down; the **style panel** for the selected element (colors and more).
- **Insert**: heading, paragraph, table, image (by link), more.
- **Image actions** (with an image selected): crop, **remove background**, replace, lock aspect ratio.
- **Element actions** (with an element selected in the preview inspector): delete, duplicate, move up/down.
- **AI button**: opens the AI panel; ask about the selected element directly.

## Insert skeleton

For a blank page, **Insert ▸ Insert skeleton** writes a minimal standards-mode document:

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8">
    <title></title>
  </head>
  <body></body>
</html>
```

Each part is there for a reason, which is why it is a command rather than something to type:

- the **doctype**, or the preview runs in quirks mode, where box sizing and table layout follow different rules than you expect;
- the **`lang`**, or a screen reader has no language to read the page in and the browser picks a font and a spell-checker for the wrong one;
- the **charset**, or a page of non-Latin text can decode as mojibake.

A viewport meta is deliberately absent: this renders in a desktop pane with no mobile viewport for it to affect.

The `lang` follows the app's UI language, so the skeleton you insert is the one your tooling is already set up for. Edit it freely afterwards.

The item only appears in edit mode, and only while the document is empty — there is nothing to insert a skeleton *into* once there is content.

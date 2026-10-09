# Keyboard shortcuts

Every editor has its own key bindings, and only a few of them work everywhere. This page is the full list, grouped by application, with the macOS and the Windows / Linux chord side by side.

## Read this first

**Almost nothing here is global.** The three access keys below are; everything after them belongs to the application that is in front of you, and changes when you switch tabs.

| Action                                          | macOS | Windows / Linux |
| ----------------------------------------------- | ----- | --------------- |
| Open the User Guide — works anywhere in the app | F1    | F1              |
| Collapse / expand the ribbon — every editor     | ⌥⌘R   | Ctrl+F1         |
| Open the Keyboard Shortcuts sheet — Word only   | ⌘/    | Ctrl+/          |

- **F1 opens this manual**, not the shortcut sheet. It belongs to the application menu, so it works from Word, Sheets, Slides, PDF or the Home tab alike.
- **The ribbon toggle is `Ctrl+F1`, or `⌥⌘R` on macOS.** Pressing bare `F1` does nothing — F1 is the manual, and the ribbon key deliberately does not share it. `Ctrl+F1` is accepted on macOS as well, though `⌥⌘R` is the key Office for Mac uses.
- **The Keyboard Shortcuts sheet is Word's alone** (`⌘/`). It lists exactly the Word table below. No other editor has one.

![Word's own Keyboard Shortcuts sheet: the action on the left, the chord on the right, grouped by File, Edit, Text, Paragraph, Insert, Review and View](img/word-shortcuts.png)

- **`⌘O`, `⌘S` and `⌘P` are implemented separately behind each editor**, and they do not quite behave alike. Word's `⌘P` first lays out a print preview so one printed sheet equals one on-screen page; PDF's opens the PDF print dialog; Sheets and Slides have their own again.
- **The rest is per-application.** When the Word tab is active, `⌘B` is Word's Bold. When Slides is active, `⌘B` is Slides' Bold, and the menu it belongs to is not the one you are used to. Where a binding only fires with something selected, or in a particular pane, the Action column says so.

## Conventions

- The macOS column uses the familiar symbols: `⌘` Command, `⌥` Option, `⇧` Shift, `⌃` Control.
- The Windows / Linux column spells them out: `Ctrl`, `Alt`, `Shift`.
- Where GenOffice genuinely needs different keys on the two platforms, the table shows it. The common cases are `⌃H` for Find & Replace (macOS owns `⌘H` for Hide), `⌥⌘G` for Go To, and Slides' close key.
- A chord that only works with a selection is written in the Action column rather than given its own column, so a table stays three columns wide.

## Word

Word has a full menu bar, so the file and formatting chords below work whenever a Word tab has focus, whichever document is open. The editing chords need a loaded, editable document with the caret in the text; they do not fire while you are typing in the Find box, the AI prompt or any other field.

### File

| Action      | macOS | Windows / Linux |
| ----------- | ----- | --------------- |
| New         | ⌘N    | Ctrl+N          |
| New Window  | ⇧⌘N   | Ctrl+Shift+N    |
| Open…       | ⌘O    | Ctrl+O          |
| Save        | ⌘S    | Ctrl+S          |
| Save As…    | ⇧⌘S   | Ctrl+Shift+S    |
| Print       | ⌘P    | Ctrl+P          |
| Close (tab) | ⌘W    | Ctrl+W          |

⌘P is not the browser's print. Word lays out a print preview first, so one printed sheet is exactly one on-screen page, then opens its own print dialog.

### Edit

| Action              | macOS | Windows / Linux |
| ------------------- | ----- | --------------- |
| Undo                | ⌘Z    | Ctrl+Z          |
| Redo                | ⇧⌘Z   | Ctrl+Shift+Z    |
| Cut                 | ⌘X    | Ctrl+X          |
| Copy                | ⌘C    | Ctrl+C          |
| Paste               | ⌘V    | Ctrl+V          |
| Paste as Plain Text | ⌥⇧⌘V  | Ctrl+Shift+V    |
| Select All          | ⌘A    | Ctrl+A          |
| Find                | ⌘F    | Ctrl+F          |
| Replace             | ⌃H    | Ctrl+H          |
| Go To               | ⌥⌘G   | Ctrl+G          |
| Keyboard Shortcuts  | ⌘/    | Ctrl+/          |

Replace is Ctrl+H on both platforms because ⌘H hides the window on macOS and never reaches Word. Go To follows Word's own split: ⌥⌘G on the Mac so ⇧⌘G stays free for Word Count.

### Text Formatting

| Action                                     | macOS    | Windows / Linux       |
| ------------------------------------------ | -------- | --------------------- |
| Bold                                       | ⌘B       | Ctrl+B                |
| Italic                                     | ⌘I       | Ctrl+I                |
| Underline                                  | ⌘U       | Ctrl+U                |
| Font… (dialog)                             | ⌘D       | Ctrl+D                |
| Increase Font Size                         | ⇧⌘.      | Ctrl+Shift+.          |
| Decrease Font Size                         | ⇧⌘,      | Ctrl+Shift+,          |
| Increase Font Size 1 pt                    | ⌘]       | Ctrl+]                |
| Decrease Font Size 1 pt                    | ⌘[       | Ctrl+[                |
| Superscript                                | ⇧⌘= / ⌘. | Ctrl+Shift+= / Ctrl+. |
| Subscript                                  | ⌘,       | Ctrl+,                |
| Change Case                                | ⇧F3      | Shift+F3              |
| Copy Formatting (with text selected)       | ⇧⌘C      | Ctrl+Shift+C          |
| Paste Formatting (after a Copy Formatting) | ⇧⌘V      | Ctrl+Alt+Shift+V      |
| Clear All Formatting                       | ⌃␣       | Ctrl+Space            |

The two formatting-painter chords are the one place the Windows key is not a plain rename: Paste Formatting carries Alt because Ctrl+Shift+V already pastes and matches style there.

### Paragraph Formatting

| Action                                      | macOS | Windows / Linux |
| ------------------------------------------- | ----- | --------------- |
| Align Left                                  | ⌘L    | Ctrl+L          |
| Center                                      | ⌘E    | Ctrl+E          |
| Align Right                                 | ⌘R    | Ctrl+R          |
| Justify                                     | ⌘J    | Ctrl+J          |
| Line Spacing 1.0                            | ⌘1    | Ctrl+1          |
| Line Spacing 1.5                            | ⌘5    | Ctrl+5          |
| Line Spacing 2.0                            | ⌘2    | Ctrl+2          |
| Increase Indent                             | ⌃M    | Ctrl+M          |
| Decrease Indent                             | ⌃⇧M   | Ctrl+Shift+M    |
| Hanging Indent                              | ⌘T    | Ctrl+T          |
| Remove Hanging Indent                       | ⇧⌘T   | Ctrl+Shift+T    |
| Clear Paragraph Formatting                  | ⌃Q    | Ctrl+Q          |
| Normal (style)                              | ⌥⌘0   | Ctrl+Alt+0      |
| Heading 1 (style)                           | ⌥⌘1   | Ctrl+Alt+1      |
| Heading 2 (style)                           | ⌥⌘2   | Ctrl+Alt+2      |
| Heading 3 (style)                           | ⌥⌘3   | Ctrl+Alt+3      |
| Paragraph… (dialog)                         | ⌥⌘M   | Ctrl+Alt+M      |
| Move Paragraph Up                           | ⌥⇧↑   | Alt+Shift+↑     |
| Move Paragraph Down                         | ⌥⇧↓   | Alt+Shift+↓     |
| Demote List Item (at the start of an item)  | Tab   | Tab             |
| Promote List Item (at the start of an item) | ⇧Tab  | Shift+Tab       |

Indent, hanging indent and Clear Paragraph Formatting are Ctrl-only on both platforms: ⌘M minimizes the window and ⌘Q quits, so a Command chord would be a silent edit on the way out. Tab only changes a list level when the caret is at the start of an item or spans several items; inside the text it inserts a tab character, and ⇧Tab always promotes.

### Insert

| Action              | macOS | Windows / Linux  |
| ------------------- | ----- | ---------------- |
| Page Break          | ⌘⏎    | Ctrl+Enter       |
| Column Break        | ⇧⌘⏎   | Ctrl+Shift+Enter |
| Line Break          | ⇧⏎    | Shift+Enter      |
| Link…               | ⌘K    | Ctrl+K           |
| New Comment         | ⌥⌘A   | Ctrl+Alt+A       |
| Insert Footnote     | ⌥⌘F   | Ctrl+Alt+F       |
| Insert Endnote      | ⌥⌘E   | Ctrl+Alt+D       |
| Date (field)        | ⌥⇧D   | Alt+Shift+D      |
| Time (field)        | ⌥⇧T   | Alt+Shift+T      |
| Non-breaking Space  | ⇧⌘␣   | Ctrl+Shift+Space |
| Non-breaking Hyphen | ⇧⌘-   | Ctrl+Shift+-     |

Insert Endnote is a real platform split, not a rename: Word uses ⌥⌘E on the Mac because ⌥⌘D belongs to the Dock there.

### Review & Tools

| Action             | macOS | Windows / Linux |
| ------------------ | ----- | --------------- |
| Track Changes      | ⇧⌘E   | Ctrl+Shift+E    |
| Word Count         | ⇧⌘G   | Ctrl+Shift+G    |
| Proofread          | F7    | F7              |
| Update Field       | F9    | F9              |
| Toggle Field Codes | ⌥F9   | Alt+F9          |

F7 runs the AI proofread and F9 updates fields; both need a loaded document. Once an editing restriction turns Track Changes on, it cannot be switched back off.

### View

| Action                     | macOS | Windows / Linux |
| -------------------------- | ----- | --------------- |
| Zoom In                    | ⌘=    | Ctrl+=          |
| Zoom Out                   | ⌘-    | Ctrl+-          |
| Zoom to 100%               | ⌘0    | Ctrl+0          |
| Show/hide formatting marks | ⌘8    | Ctrl+Shift+8    |

The formatting-marks chord follows Word's own split, ⌘8 against Ctrl+Shift+8. Word takes zoom only through the three menu chords above — ⌘+ does nothing here.

F1 opens the **User Guide**, not this list. The shortcut sheet is ⌘/ or Ctrl+/, and it is a Word menu item, so it only appears while a Word tab is active.

## Sheets

GenOffice Sheets is an Excel-compatible grid. Most of the shortcuts below are the Excel ones you already know.

> **Editing wins.** Almost every shortcut in Sheets is ignored while you are editing a cell or typing in the **formula bar**, so that typing, moving the caret and Backspace keep their normal text-editing meaning. If a key seems to do nothing, check whether you are mid-edit: press Esc or Enter first, then try again. Commands that act on the selection also do nothing while a dialog is open. Where a chord is an exception to this rule, the table says so.

Two conventions are specific to Sheets. `⌘` and `⌃` are the same modifier slot — both are accepted on every platform, so `⌘1` and `Ctrl+1` are the same chord. **An empty Windows / Linux cell means the chord exists on macOS only.**

### Navigation

| Action                                             | macOS      | Windows / Linux |
| -------------------------------------------------- | ---------- | --------------- |
| Next worksheet                                     | ⌘PgDn      | Ctrl+PgDn       |
| Previous worksheet                                 | ⌘PgUp      | Ctrl+PgUp       |
| Next worksheet (Excel for Mac parity)              | ⌥→         |                 |
| Previous worksheet (Excel for Mac parity)          | ⌥←         |                 |
| Jump to the first unfrozen cell                    | ⌘Home      | Ctrl+Home       |
| Jump to the last used cell (content or formatting) | ⌘End       | Ctrl+End        |
| Go To dialog                                       | ⌘G, F5     | Ctrl+G, F5      |
| Insert a new worksheet                             | ⇧F11       | Shift+F11       |
| Move the active cell one screen down (and scroll)  | Page Down  | Page Down       |
| Move the active cell one screen up (and scroll)    | Page Up    | Page Up         |
| Move the active cell one screen right (and scroll) | ⌥Page Down | Alt+Page Down   |
| Move the active cell one screen left (and scroll)  | ⌥Page Up   | Alt+Page Up     |

### Selection

| Action                                                                            | macOS           | Windows / Linux                                     |
| --------------------------------------------------------------------------------- | --------------- | --------------------------------------------------- |
| Select the whole column of the active cell                                        | ⌃Space          | Ctrl+Space                                          |
| Select the whole row of the active cell                                           | ⇧Space          | Shift+Space                                         |
| Jump to the edge of the data block — any of the four arrows                       | ⌘↑ ⌘↓ ⌘← ⌘→     | Ctrl+↑ Ctrl+↓ Ctrl+← Ctrl+→                         |
| Extend the selection to the edge of the data block; press again to shrink it back | ⌘⇧↑ ⌘⇧↓ ⌘⇧← ⌘⇧→ | Ctrl+Shift+↑ Ctrl+Shift+↓ Ctrl+Shift+← Ctrl+Shift+→ |
| Jump to the first column of the current row                                       | Home            | Home                                                |

### Editing

| Action                                                                                               | macOS                       | Windows / Linux     |
| ---------------------------------------------------------------------------------------------------- | --------------------------- | ------------------- |
| Fill the whole selection with the current entry                                                      | ⌃⏎ or ⌘⏎                    | Ctrl+Enter          |
| Start a new line inside the cell                                                                     | ⌥⏎, or ⌃⌥⏎ / ⌘⌥⏎ on the Mac | Alt+Enter           |
| Edit the active cell in place                                                                        | F2                          | F2                  |
| Bold                                                                                                 | ⌘B                          | Ctrl+B              |
| Italic                                                                                               | ⌘I                          | Ctrl+I              |
| Underline                                                                                            | ⌘U                          | Ctrl+U              |
| Find                                                                                                 | ⌘F                          | Ctrl+F              |
| Fill right from the leftmost selected column                                                         | ⌘R                          | Ctrl+R              |
| Cycle a formula reference between relative and absolute                                              | F4                          | F4                  |
| Cycle a formula reference between relative and absolute (Excel for Mac parity, while editing a cell) | ⌘T                          |                     |
| Copy                                                                                                 | ⌘C                          | Ctrl+C              |
| Cut                                                                                                  | ⌘X                          | Ctrl+X              |
| Paste                                                                                                | ⌘V                          | Ctrl+V              |
| Paste values only                                                                                    | ⌘⇧V                         | Ctrl+Shift+V        |
| Flash Fill from the pattern above                                                                    | ⌘E                          | Ctrl+E              |
| Undo                                                                                                 | ⌘Z                          | Ctrl+Z              |
| Redo                                                                                                 | ⇧⌘Z                         | Ctrl+Shift+Z        |
| Select the whole sheet                                                                               | ⌘A                          | Ctrl+A              |
| Insert today's date                                                                                  | ⌘;                          | Ctrl+;              |
| Insert the current time                                                                              | ⌘⇧;                         | Ctrl+Shift+;        |
| Clear the contents of every selected cell                                                            | ⌫ or ⌦                      | Backspace or Delete |

AutoSum is macOS only. Press ⌘⇧T.

### Formatting

| Action                                                           | macOS | Windows / Linux |
| ---------------------------------------------------------------- | ----- | --------------- |
| Format Cells dialog                                              | ⌘1    | Ctrl+1          |
| Strikethrough (with cells selected)                              | ⌘5    | Ctrl+5          |
| Increase font size (with cells selected)                         | ⌘⇧.   | Ctrl+Shift+.    |
| Decrease font size (with cells selected)                         | ⌘⇧,   | Ctrl+Shift+,    |
| Number format: General                                           | ⌘⇧`   | Ctrl+Shift+`    |
| Number format: number with two decimals                          | ⌘⇧1   | Ctrl+Shift+1    |
| Number format: time `h:mm AM/PM`                                 | ⌘⇧2   | Ctrl+Shift+2    |
| Number format: date `d-mmm-yy`                                   | ⌘⇧3   | Ctrl+Shift+3    |
| Number format: currency                                          | ⌘⇧4   | Ctrl+Shift+4    |
| Number format: percent                                           | ⌘⇧5   | Ctrl+Shift+5    |
| Number format: scientific                                        | ⌘⇧6   | Ctrl+Shift+6    |
| All borders (with cells selected)                                | ⌘⇧7   | Ctrl+Shift+7    |
| No borders (with cells selected)                                 | ⌘⇧-   | Ctrl+Shift+-    |
| Cancel Format Painter (only while Format Painter is switched on) | Esc   | Esc             |

### Insert & delete

| Action                                                                                      | macOS                 | Windows / Linux          |
| ------------------------------------------------------------------------------------------- | --------------------- | ------------------------ |
| Insert cells — whole rows or columns insert directly, anything else opens the Insert dialog | ⌘⇧=                   | Ctrl+Shift+=             |
| Insert cells — numeric keypad `+`, same dialog                                              | ⌘ plus the keypad `+` | Ctrl plus the keypad `+` |
| Delete cells — whole rows or columns delete directly, anything else opens the Delete dialog | ⌘-                    | Ctrl+-                   |
| Delete cells — numeric keypad `-`, same dialog                                              | ⌘ plus the keypad `-` | Ctrl plus the keypad `-` |
| Hide the selected rows                                                                      | ⌘9                    | Ctrl+9                   |
| Unhide the hidden rows in the selection                                                     | ⌘⇧9                   | Ctrl+Shift+9             |
| Hide the selected columns                                                                   | ⌘0                    | Ctrl+0                   |
| Unhide the hidden columns in the selection                                                  | ⌘⇧0                   | Ctrl+Shift+0             |

### Data & review

| Action                                         | macOS | Windows / Linux |
| ---------------------------------------------- | ----- | --------------- |
| Insert Function dialog                         | ⇧F3   | Shift+F3        |
| Name Manager                                   | ⌃F3   | Ctrl+F3         |
| Add or edit a note on the active cell          | ⇧F2   | Shift+F2        |
| Insert hyperlink                               | ⌘K    | Ctrl+K          |
| Trace precedents (cell must contain a formula) | ⌘[    | Ctrl+[          |
| Trace dependents                               | ⌘]    | Ctrl+]          |
| Group the selected rows                        | ⌥⇧→   | Alt+Shift+→     |
| Ungroup the selected rows                      | ⌥⇧←   | Alt+Shift+←     |
| Recalculate the whole workbook                 | F9    | F9              |
| Recalculate the active sheet only              | ⇧F9   | Shift+F9        |
| Toggle the filter on the selected range        | ⌘⇧F   |                 |

Excel's filter shortcut, Ctrl+Shift+L, does nothing on Windows or Linux. Press ⌘⇧F on macOS, or use the filter command on the ribbon on any platform.

### Window

| Action                                                      | macOS      | Windows / Linux |
| ----------------------------------------------------------- | ---------- | --------------- |
| Save the workbook                                           | ⌘S         | Ctrl+S          |
| Save As                                                     | ⇧⌘S        | Ctrl+Shift+S    |
| Open a workbook                                             | ⌘O         | Ctrl+O          |
| Print                                                       | ⌘P         | Ctrl+P          |
| Close the workbook (macOS) / Quit the app (Windows & Linux) | ⌘W         | Ctrl+Q          |
| Open the User Guide                                         | F1         | F1              |
| Zoom in                                                     | ⌘=         | Ctrl+=          |
| Zoom with the mouse wheel                                   | ⌘ + scroll | Ctrl + scroll   |
| Show or hide formulas in the cells                          | ⌘`         | Ctrl+`          |
| Expand or collapse the formula bar                          | ⌘⇧U        | Ctrl+Shift+U    |

## Slides

GenOffice Slides is a PowerPoint-style presentation editor. Most keys do the same thing wherever you are, but a few change meaning with focus — inside a text box, or while the **thumbnail pane**, **outline pane** or **slide sorter** has keyboard focus. Those cases are marked in the Action column.

### File and app

| Action                 | macOS | Windows / Linux |
| ---------------------- | ----- | --------------- |
| Save                   | ⌘S    | Ctrl+S          |
| Save as                | ⇧⌘S   | Ctrl+Shift+S    |
| Open                   | ⌘O    | Ctrl+O          |
| Close current tab      | ⌘W    | Ctrl+Q          |
| Print                  | ⌘P    | Ctrl+P          |
| Undo                   | ⌘Z    | Ctrl+Z          |
| Redo                   | ⇧⌘Z   | Ctrl+Shift+Z    |
| Redo (alternate chord) | ⌘Y    | Ctrl+Y          |

Closing a tab is the one place the two platforms split completely: ⌘W on macOS, and on Windows and Linux Ctrl+Q, which quits the whole app rather than the tab.

### Editing

| Action                                                | macOS  | Windows / Linux |
| ----------------------------------------------------- | ------ | --------------- |
| New slide                                             | ⇧⌘N    | Ctrl+M          |
| Find (not while editing text in a shape)              | ⌘F     | Ctrl+F          |
| Annotate the selection with AI (with shapes selected) | ⌘K     | Ctrl+K          |
| Zoom in                                               | ⌘=     | Ctrl+=          |
| Zoom out                                              | ⌘-     | Ctrl+-          |
| Actual size (100%)                                    | ⌘0     | Ctrl+0          |
| Select all                                            | ⌘A     | Ctrl+A          |
| Cycle shape selection forward / back                  | ⇥ / ⇧⇥ | Tab / Shift+Tab |

New slide is the other platform split. Each platform has its own key and the other one does nothing: ⇧⌘N on macOS, Ctrl+M on Windows and Linux. ⌘M cannot do the job on macOS because that is the system Minimize key.

Select all follows focus: with the thumbnail pane, outline pane or slide sorter focused it selects every slide, and anywhere else it selects every element on the current slide. Tab cycles through the shapes on the slide, and only when the slide canvas itself is focused.

### Text formatting

These act on the text of a selected shape, or on the caret's paragraph when you are inside the text box.

| Action                                     | macOS | Windows / Linux |
| ------------------------------------------ | ----- | --------------- |
| Bold                                       | ⌘B    | Ctrl+B          |
| Italic                                     | ⌘I    | Ctrl+I          |
| Underline                                  | ⌘U    | Ctrl+U          |
| Align left                                 | ⌘L    | Ctrl+L          |
| Align centre                               | ⌘E    | Ctrl+E          |
| Align right                                | ⌘R    | Ctrl+R          |
| Justify                                    | ⌘J    | Ctrl+J          |
| Increase font size by 1 pt                 | ⌘]    | Ctrl+]          |
| Decrease font size by 1 pt                 | ⌘[    | Ctrl+[          |
| Grow font to next size on the ladder       | ⇧⌘>   | Ctrl+Shift+>    |
| Shrink font to previous size on the ladder | ⇧⌘<   | Ctrl+Shift+<    |

The font-ladder keys also take the . and , keys with Shift held, for the keyboard layouts that report those instead of the punctuation keys.

### Clipboard and objects

| Action                                      | macOS       | Windows / Linux    |
| ------------------------------------------- | ----------- | ------------------ |
| Copy                                        | ⌘C          | Ctrl+C             |
| Cut                                         | ⌘X          | Ctrl+X             |
| Paste                                       | ⌘V          | Ctrl+V             |
| Paste format only                           | ⇧⌘V         | Ctrl+Shift+V       |
| Copy format only (with shapes selected)     | ⇧⌘C         | Ctrl+Shift+C       |
| Duplicate in place (with shapes selected)   | ⌘D          | Ctrl+D             |
| Group selection (with shapes selected)      | ⌘G          | Ctrl+G             |
| Ungroup selection (with shapes selected)    | ⇧⌘G         | Ctrl+Shift+G       |
| Delete selection (with shapes selected)     | ⌫ / Delete  | Backspace / Delete |
| Delete the selected vertex (in Edit Points) | ⌫ / Delete  | Backspace / Delete |
| Nudge selection (with shapes selected)      | ↑ ↓ ← →     | Arrow keys         |
| Nudge one screen pixel                      | ⌘↑ ⌘↓ ⌘← ⌘→ | Ctrl+Arrow keys    |
| Resize about the centre                     | ⇧↑ ⇧↓ ⇧← ⇧→ | Shift+Arrow keys   |

Inside a text box, copy, cut and paste act on the text rather than on the shape. Shift with the arrow keys resizes each shape about its own centre, leaves connectors alone, and commits a multi-selection as one undo step.

### Slides — thumbnail pane, outline pane or sorter focused

Slide-level commands of the PowerPoint kind only fire while one of those three panes has keyboard focus.

| Action                                 | macOS                          | Windows / Linux                                |
| -------------------------------------- | ------------------------------ | ---------------------------------------------- |
| Previous / next slide                  | ↑ ↓ (also Page Up / Page Down) | ArrowUp / ArrowDown (also Page Up / Page Down) |
| First / last slide                     | Home / End                     | Home / End                                     |
| Extend selection to first / last slide | ⇧Home / ⇧End                   | Shift+Home / Shift+End                         |
| Select all slides                      | ⌘A                             | Ctrl+A                                         |
| Delete selected slides                 | ⌫ / Delete                     | Backspace / Delete                             |
| Copy slides                            | ⌘C                             | Ctrl+C                                         |
| Cut slides                             | ⌘X                             | Ctrl+X                                         |

Moving between slides with the arrow keys is the exception — it works from the canvas too, with nothing selected. Deleting slides does not apply while a pen tool is active or in **Reading view**.

### Inside a text box

| Action                                                   | macOS                             | Windows / Linux                       |
| -------------------------------------------------------- | --------------------------------- | ------------------------------------- |
| Enter text editing on a selected text shape              | any printable character, ⏎, or F2 | any printable character, Enter, or F2 |
| Commit the text and reselect the shape                   | Esc or F2                         | Esc or F2                             |
| Commit the text and reselect the shape (alternate chord) | ⌘⏎                                | Ctrl+⏎                                |
| Jump to the next placeholder                             | ⌃⏎                                | Ctrl+⏎                                |
| Demote / promote list level                              | ⇥ / ⇧⇥                            | Tab / Shift+Tab                       |
| Next / previous table cell                               | ⇥ / ⇧⇥                            | Tab / Shift+Tab                       |
| Soft line break                                          | ⇧⏎                                | Shift+⏎                               |

Tab means two different things: the next cell inside a table, and demote the list level inside a text shape.

### Slide show

| Action                       | macOS             | Windows / Linux                      |
| ---------------------------- | ----------------- | ------------------------------------ |
| Start from the beginning     | F5                | F5                                   |
| Start from the current slide | ⇧F5 or ⌘⏎         | ⇧F5                                  |
| Next slide                   | → ↓ ␣ ⏎ Page Down | Right, Down, Space, Enter, Page Down |
| Previous slide               | ← ↑ Page Up       | Left, Up, Page Up                    |
| First slide                  | Home              | Home                                 |
| Last slide                   | End               | End                                  |
| Go to slide N                | type N, then ⏎    | type N, then Enter                   |
| Black screen                 | B or .            | B or .                               |
| White screen                 | W or ,            | W or ,                               |
| Restore after a blackout     | any key           | any key                              |
| End the show                 | Esc               | Esc                                  |

Starting from the current slide takes ⌘⏎ as well as ⇧F5 on macOS. On Windows the same key stays reserved for moving between placeholders.

### Reading view

| Action            | macOS             | Windows / Linux                      |
| ----------------- | ----------------- | ------------------------------------ |
| Next slide        | → ↓ ␣ ⏎ Page Down | Right, Down, Space, Enter, Page Down |
| Previous slide    | ← ↑ Page Up       | Left, Up, Page Up                    |
| Exit reading view | Esc               | Esc                                  |

Press **Esc** to back out of whatever you are in — a crop, a shape you are drawing, **Edit Points**, a group, the pen, highlighter or eraser tool, the media player or **Reading view** — and to clear the selection.

Two more are modifiers rather than chords: hold **Shift** while drawing to keep the new shape's proportions, and hold **Shift** while rotating to snap to 15° steps.

## PDF

PDF has no menu of its own, so every chord below follows the active tab and works only while a PDF tab is in front.

### File and Editing

| Action             | macOS | Windows / Linux |
| ------------------ | ----- | --------------- |
| Save               | ⌘S    | Ctrl+S          |
| Undo               | ⌘Z    | Ctrl+Z          |
| Redo               | ⇧⌘Z   | Ctrl+Shift+Z    |
| Find (open search) | ⌘F    | Ctrl+F          |
| Print              | ⌘P    | Ctrl+P          |

Undo and Redo step the document history; when the cursor sits in a text field they leave ⌘Z to that field. ⌘P opens the same in-app print dialog as the toolbar's Print button.

### View

| Action            | macOS            | Windows / Linux     |
| ----------------- | ---------------- | ------------------- |
| Zoom In           | ⌘=               | Ctrl+=              |
| Zoom Out          | ⌘-               | Ctrl+-              |
| Fit to width      | ⌘0               | Ctrl+0              |
| Zoom (continuous) | ⌘ + scroll wheel | Ctrl + scroll wheel |

Zoom In also answers to ⌘+ and Ctrl++.

### Page Navigation

| Action                 | macOS             | Windows / Linux   |
| ---------------------- | ----------------- | ----------------- |
| Next page              | →                 | →                 |
| Previous page          | ←                 | ←                 |
| Scroll down one screen | PageDown or Space | PageDown or Space |
| Scroll up one screen   | PageUp            | PageUp            |
| Scroll down            | ↓                 | ↓                 |
| Scroll up              | ↑                 | ↑                 |
| Jump to first page     | Home              | Home              |
| Jump to last page      | End               | End               |

These take no modifier, so both platforms read the same. ↓ and ↑ are the one row that changes with focus: in the document they nudge the page, but with focus in the thumbnail sidebar they step a whole page. They do nothing while you are typing in a field.

### Selection and Editing

| Action                     | macOS               | Windows / Linux     |
| -------------------------- | ------------------- | ------------------- |
| Delete selected annotation | Delete or Backspace | Delete or Backspace |

Only active when an annotation, stamp, redaction or image is selected.

### Escape and Dialogs

| Action                   | macOS  | Windows / Linux |
| ------------------------ | ------ | --------------- |
| Cancel the topmost state | Escape | Escape          |
| Confirm the open dialog  | Enter  | Enter           |

Escape unwinds one layer at a time — an open dialog first, then the AI popover, a text draft, the search panel and finally the selection itself. With the cursor in the search box, Enter and ⇧Enter step to the next and previous match, and Escape closes the panel; with the cursor in a new text note, ⌘⏎ commits it.

## Markdown

GenOffice Markdown edits Markdown with a live preview alongside.

### File and view

| Action                  | macOS    | Windows / Linux  |
| ----------------------- | -------- | ---------------- |
| Save                    | ⌘S       | Ctrl+S           |
| Save as                 | ⇧⌘S      | Ctrl+Shift+S     |
| Print                   | ⌘P       | Ctrl+P           |
| Find                    | ⌘F       | Ctrl+F           |
| Toggle source / preview | ⌘E       | Ctrl+E           |
| Zoom in                 | ⌘= or ⌘+ | Ctrl+= or Ctrl++ |
| Zoom out                | ⌘-       | Ctrl+-           |
| Actual size (100%)      | ⌘0       | Ctrl+0           |

Zoom out also takes ⌘_ and Ctrl+_. Pinching the trackpad, or holding ⌘ or Ctrl and scrolling over the document, zooms too.

### Blocks

| Action                                             | macOS | Windows / Linux      |
| -------------------------------------------------- | ----- | -------------------- |
| Duplicate the selected blocks                      | ⌘D    | Ctrl+D               |
| Delete the selected blocks                         | ⇧⌘⌫   | Ctrl+Shift+Backspace |
| Move the selected block up                         | ⇧⌘↑   | Ctrl+Shift+ArrowUp   |
| Move the selected block down                       | ⇧⌘↓   | Ctrl+Shift+ArrowDown |
| Insert a link on the selection                     | ⌘K    | Ctrl+K               |
| Remove the link (press again with the link active) | ⌘K    | Ctrl+K               |

### Inside a math block

These apply only while the floating LaTeX editor is open.

| Action                              | macOS | Windows / Linux |
| ----------------------------------- | ----- | --------------- |
| Apply the formula                   | ⌘⏎    | Ctrl+⏎          |
| Dismiss the editor without applying | Esc   | Esc             |

Applying an empty formula deletes the block instead. The editor also closes on any change to the document, and clicking outside dismisses it.

## HTML

GenOffice HTML shows a code editor beside a live preview. The keys that act on a page element apply only when the preview itself has keyboard focus.

### File and source view

| Action                              | macOS    | Windows / Linux  |
| ----------------------------------- | -------- | ---------------- |
| Save                                | ⌘S       | Ctrl+S           |
| Save as                             | ⇧⌘S      | Ctrl+Shift+S     |
| Cycle view (edit / split / preview) | ⌘\       | Ctrl+\           |
| Find                                | ⌘F       | Ctrl+F           |
| Zoom in                             | ⌘= or ⌘+ | Ctrl+= or Ctrl++ |
| Zoom out                            | ⌘-       | Ctrl+-           |
| Actual size (100%)                  | ⌘0       | Ctrl+0           |

Zoom out also takes ⌘_ and Ctrl+_.

### Formatting and history

| Action                 | macOS | Windows / Linux |
| ---------------------- | ----- | --------------- |
| Bold                   | ⌘B    | Ctrl+B          |
| Italic                 | ⌘I    | Ctrl+I          |
| Undo                   | ⌘Z    | Ctrl+Z          |
| Redo                   | ⇧⌘Z   | Ctrl+Shift+Z    |
| Redo (alternate chord) | ⌘Y    | Ctrl+Y          |

Inside the code editor or a form field, the field's own editing keys take over instead.

### In the preview frame, with an element selected

| Action                            | macOS      | Windows / Linux    |
| --------------------------------- | ---------- | ------------------ |
| Ask AI about the element          | ⌘K         | Ctrl+K             |
| Delete the element                | ⌫ / Delete | Backspace / Delete |
| Previous sibling                  | ↑          | ArrowUp            |
| Next sibling                      | ↓          | ArrowDown          |
| Select parent                     | ←          | ArrowLeft          |
| Select first child                | →          | ArrowRight         |
| Move element up in the document   | ⌥↑         | Alt+ArrowUp        |
| Move element down in the document | ⌥↓         | Alt+ArrowDown      |
| Deselect                          | Esc        | Esc                |

Each of these needs an element to be selected already; with nothing selected the key does nothing.

### Inside the preview frame, editing text in place

| Action                            | macOS | Windows / Linux |
| --------------------------------- | ----- | --------------- |
| Commit the edit                   | Esc   | Esc             |
| Commit the edit (alternate chord) | ⌘⏎    | Ctrl+⏎          |
| Bold the selected range           | ⌘B    | Ctrl+B          |
| Italic the selected range         | ⌘I    | Ctrl+I          |

Committing keeps what you typed and returns the element to the selected state.

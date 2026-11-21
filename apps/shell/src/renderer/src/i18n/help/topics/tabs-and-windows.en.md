# Tabs and window management

In GenOffice, every file you open—whether it is a Word document, an Excel workbook, a PowerPoint presentation, or a PDF—lives inside a single unified window by default, with a browser-style tab strip running across the top. You can jump between different document types just as you would switch tabs in Chrome, without hunting through a dozen separate windows in your dock or taskbar.

![The tab strip: Home plus two documents](img/tabs.png)

## Basics

- **Switch between tabs**: Click any tab in the top strip to bring that document to the front, or hover your pointer anywhere over the tab bar and roll the mouse wheel to cycle smoothly through your open files in order.
- **Close a tab**: Click the `×` button on the right side of a tab, or press `⌘W` on macOS (`Ctrl+W` on Windows and Linux). If the document has unsaved edits, GenOffice always pauses and prompts you to save, discard your changes, or cancel before closing anything.
- **Open a new tab**: Click the `+` button at the far right end of the tab strip to open a fresh blank document tab and jump straight into its editor.
- **Immediate tab activation**: Clicking a tab switches to that document the instant your mouse button presses down rather than waiting for the release, which makes flipping back and forth to compare two files feel instantaneous.
- **Rename an open file**: When you want to rename a saved file on disk without leaving your workspace, either double-click its tab title or right-click the tab and choose **Rename** from the context menu; a normal single click only activates the tab so you never enter rename mode by accident.

## Rename: double-click or right-click a tab

On any document tab that already corresponds to a saved file on disk, **double-clicking the tab** (or choosing **Rename** from the tab's right-click menu) turns the title straight into an inline text field prefilled with the current file name. Type the new name and press `Enter`—or simply click away so the field loses focus—to commit the change, or press `Escape` at any point to cancel and keep the original name. If you use an input method editor (IME) for Chinese, Japanese, or Korean, pressing `Enter` to confirm a candidate word is handled properly and will never prematurely submit half a file name. The rename runs through the exact same safety checks as renaming a row on the Home screen: if you omit the extension, the existing extension is preserved automatically, while illegal path characters, reserved system names, or name collisions in the same folder are caught with a clear warning. Once confirmed, the file is renamed on disk immediately and both your Recent and Starred lists update in place.

## Drag to reorder

To organize your workspace when many files are open, press and hold any document tab and drag it left or right along the strip. Neighboring tabs slide out of the way in real time to show where the tab will land, and the new order locks in as soon as you release the mouse button. A built-in 4-pixel movement threshold ensures that ordinary clicks never cause tabs to jitter or shift unintentionally.

## Context menu

Right-clicking any document tab opens a context menu that covers both window management for the tab itself and file operations for the document behind it. When the tab is backed by a saved file on disk, the top line of the menu displays a shortened read-only path so you can immediately tell which folder a file belongs to when you have two similarly named documents open at once.

From the upper section of the menu, you can trigger **Save**, **Save As…**, or **Print…** directly without switching into the editor's ribbon first, as well as **Rename** (which opens the same inline tab editor as double-clicking), **Duplicate**, **Copy Path**, **Copy File Name**, **Reveal in Finder / File Explorer**, or **Open in Default App** (on an unsaved untitled tab, file-backed actions remain disabled until your first save). The lower section handles layout and tab cleanup: **Split Beside** and **Open in New Window** (available on every detachable document tab), **Pin Tab** to keep a background tab's renderer alive, **Move Left / Move Right**, and **Close**, **Close Other Tabs**, **Close Tabs to the Right**, **Close All Tabs**, or **Reopen Closed Tab**, all following the active UI language.

## Tear off into a window

Whenever you want to place two documents side by side on a wide monitor or move one to a second display, simply **drag a tab vertically out of the tab strip** (or right-click the tab and choose **Open in New Window**). The tab detaches smoothly into a standalone window carrying the live document, complete with any unsaved edits and your full undo history. A detached window is a first-class GenOffice window where you can continue editing, saving, and using the AI assistant; when you are done comparing, drag its tab back over the main window's tab strip and it will dock right back into place.

## The all-tabs list

When you have enough documents open that the tab strip fills the width of the window, a compact tab-list button appears at the right end of the strip. Clicking it pops up a native system menu listing every open tab in full—because it is a native menu, it is never clipped by the editor canvas below—and you can jump to any open document with a click or with the up and down arrow keys.

## The toolbar row inside every editor

At the top of every editor tab sits a quick-access toolbar row (its exact placement and visual grouping adapt slightly to fit each editor's layout):

- **Save and Save As**: Click Save (or press `⌘S` / `Ctrl+S`) at any time to write your latest changes to disk, or use **Save As** to branch off a new copy under a different name or folder without touching the original file.
- **Multi-level Undo and Redo**: Step backward and forward through your editing session with dedicated undo and redo controls; every open tab maintains its own independent history stack so edits in one file never interfere with another.
- **Find**: Click the Find button (or press `⌘F` / `Ctrl+F`) to open that editor's search bar for jumping to matches or running find-and-replace across the document.
- **AutoSave toggle**: When AutoSave is switched on, GenOffice automatically writes your edits to disk after a brief pause in typing; when you switch it off, changes stay in memory until you save manually, and an unsaved indicator on the tab title reminds you that edits are waiting to be written.

## The Home tab

The leftmost **Home** tab is permanently pinned to the start of the tab strip and cannot be closed. Whenever you want to return to your recent files, starred documents, local folders, or the knowledge-base Q&A panel from inside an editor, simply click the Home tab at the top left or choose **File ▸ Back to Home** from the menu bar.

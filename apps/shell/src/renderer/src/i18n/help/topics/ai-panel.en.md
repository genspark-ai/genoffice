# The AI assistant panel

Every editor can summon the AI panel: select something, give an instruction, watch the streamed result.

## Opening and using it

![The AI panel in Docs](img/ai-panel.png)

- Entries: the **AI button** in each editor's ribbon, **Ask AI** in context menus, or Ask AI on the markup bar.
- Describe the task in plain language (rewrite this / make this column percentages / re-layout this page...) and press Enter.
- Replies render **streamed**; when the AI needs tools (read the document, edit it) it executes and continues until done.
- **Stop**: interrupt the current turn at any time.
- **Queueing**: while a reply is streaming, Enter queues your message instead — a strip above the input counts them; open it to edit, remove, reorder, pause or clear. The next message sends itself when the turn ends; Esc still stops the turn.

## What it can do

- **Docs**: rewrite/expand/translate/summarize, insert tables and pictures, adjust formatting; every turn snapshots first.
- **Sheets**: formulas, data fill, bulk transforms, formatting.
- **Slides**: whole-deck generation, layout adjustment, copy rewriting.
- **PDF**: Q&A and summaries against selected text or pages.
- **Markdown / HTML**: rewrite, extend, translate.
- **Home ▸ Ask AI** answers from your starred files and cites them — the knowledge-base panel has its own section on Home, and the Docs AI panel carries the same skill.

## Rollback and safety

- The Docs panel keeps a **version list**: one snapshot per turn, roll back to any of them, and the rollback itself is Ctrl+Z-able. Snapshots survive reopening the document.
- AI edits go through the same editing pipeline as manual ones (undoable, save-gated) — nothing bypasses your save confirmation.

## Hiding a selection from the AI

- In Slides, HTML and Markdown a selection can be kept out of the model's eyes: right-click it ▸ **Hide the selection from AI**, name the placeholder, and the model reads `{{name}}` in its place. (Markdown arms the item with the eye-off toggle on its ribbon; HTML's item lives in the source panel.)
- The file keeps the real text — saving, export and print are untouched, and the mark survives reopen. The same item unhides (Markdown gets a separate **Show the selection to AI again** entry).
- HTML hides element contents, attribute values and script string literals; the rendered page never changes.
- Slides also hides whole pictures, videos and audio, and a slide with hidden content is never sent as a screenshot: the automatic layout check skips the page and says so.
- This is for the in-app AI only — the CLI, MCP tools and headless export read the file directly and still see the text.

## Privacy

- Instructions and the relevant document content go to the **model service you configured** (Genspark hosted or a custom endpoint, next chapter); no configuration, no sending.
- Local files are uploaded nowhere else; BYOK keys are stored in the app's settings file on this machine and sent only in request headers.

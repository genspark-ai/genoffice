# The AI assistant panel

Every editor can summon the AI panel: select something, give an instruction, watch the streamed result.

## Opening and using it

![The AI panel in Docs](img/ai-panel.en.png)

- Entries: the **AI button** in each editor's ribbon, **Ask AI** in context menus, or Ask AI on the markup bar.
- Describe the task in plain language (rewrite this / make this column percentages / re-layout this page...) and press Enter.
- Replies render **streamed**; when the AI needs tools (read the document, edit it, run a script) it executes and continues until done.
- **Stop**: interrupt the current turn at any time.

## What it can do

- **Docs**: rewrite/expand/translate/summarize, insert tables and pictures, adjust formatting; every turn snapshots first.
- **Sheets**: formulas, data fill, bulk transforms (optionally via the run_script sandbox), formatting.
- **Slides**: whole-deck generation, layout adjustment, copy rewriting.
- **PDF**: Q&A and summaries against selected text or pages.
- **Markdown / HTML**: rewrite, extend, translate.

## Rollback and safety

- The Docs panel keeps a **version list**: one snapshot per turn, roll back to any of them, and the rollback itself is Ctrl+Z-able. Snapshots survive reopening the document.
- AI edits go through the same editing pipeline as manual ones (undoable, save-gated) — nothing bypasses your save confirmation.

## Privacy

- Instructions and the relevant document content go to the **model service you configured** (Genspark hosted or a custom endpoint, next chapter); no configuration, no sending.
- Local files are uploaded nowhere else; BYOK keys live only in request headers — never on disk or in logs.

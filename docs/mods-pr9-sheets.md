# PR 9 — sheets

Status: **implemented**. The last of the six editors, and the one whose assistant
is a different component.

## 1. What ships

The same three pieces the other editors got, adapted to sheets' own assistant:

| Piece   | What it is                                                                                      |
| ------- | ----------------------------------------------------------------------------------------------- |
| bridge  | `featuresList()` on `DesktopApi` and the preload (`window.desktopApi`)                          |
| binding | `renderer/mods/EditorSlots.tsx`, the shared renderer bound to that bridge                       |
| place   | an `editor.side` rail beside the copilot, whose press opens the copilot and hands it the prompt |

## 2. Why sheets needed its own pass

Sheets' copilot is not the shared `AiPanel`; it is a fully controlled
`AiChatPanel` (the composer text arrives as the `prompt` prop and leaves through
`onPromptChange`). So `ask` is: open the panel (`setIsCopilotOpen(true)`) and hand
it the prompt (`onPromptChange`). No new mechanism was invented for the editor.

## 3. Layout, without disturbing the sheet

`.sheet-body` is a two-column grid (copilot, sheet). A rail is only laid out while
one is actually drawn:

    .sheet-body:has(> .slot-place--editor-side) {
      grid-template-columns: auto var(--copilot-width, 360px) minmax(0, 1fr);
    }

With no feature registered, `SlotHost` renders nothing, the `:has()` rule does not
match, and the sheet body is exactly the two columns it has always been. The rail's
own look (220px, a subtle separator, its own scroll) lives in the shared
stylesheet, so every editor's rail looks the same.

## 4. Files

| File                                          | Change                                             |
| --------------------------------------------- | -------------------------------------------------- |
| apps/sheets/src/shared/desktop-api.ts         | `featuresList` on the API                          |
| apps/sheets/src/preload/index.ts              | the one-line bridge                                |
| apps/sheets/src/renderer/mods/EditorSlots.tsx | new: the binding                                   |
| apps/sheets/src/renderer/ExcelShell.tsx       | the rail beside the copilot, and the ask hand-off  |
| apps/sheets/src/renderer/styles.css           | the rail's grid column, gated on the rail existing |
| packages/ui/src/mods/mods.css                 | the shared editor-rail look                        |

## 5. Verification

- sheets typechecks with 0 errors; all six editors and the shell do;
- the full build (`npm run build:all`) succeeds, which compiles the changed
  preloads and renderers;
- agent-core (210) and shell (30) tests stay green; theme, english-only and
  prettier gates pass;
- the picture harness draws the rail through the same component.

## 6. Out of scope

- sheets has no file-opener on its bridge, so an open action draws disabled there,
  exactly as in the other four editors;
- ribbon, menu and status places: they need each app's ribbon metadata.

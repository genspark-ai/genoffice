# PR 8 — the other five editors

Status: **implemented**. Follows docs/mods-pr7-editor-slots.md, which opened the
first editor place.

## 1. What ships

The same three pieces Docs got, in the four editors that share its shape (slides,
pdf, markdown, html):

| Piece   | What it is                                                                                                                   |
| ------- | ---------------------------------------------------------------------------------------------------------------------------- |
| bridge  | `featuresList()` on the app's own API type and preload, reading the shell's features:list channel                            |
| binding | `renderer/mods/EditorSlots.tsx`: the shared renderer bound to that app's bridge                                              |
| place   | an `editor.side` rail inside the app's own AI dock, whose presses hand their prompt to that app's own assistant (its preset) |

## 2. A host that cannot serve an action does not pretend

The four editors have no file-opener on their bridge (only pdf has one), and none
of them has a panel container for a feature's own view. So the shared contract now
says so:

- `SlotActions.open` and `SlotActions.panel` are optional;
- `canRunAction(action, actions)` decides whether a host can serve a declared
  action, and the renderer draws such a control **disabled** — a button that
  cannot do its job is shown as unavailable rather than doing nothing.

The shell and Docs keep providing all three actions, so nothing in their panels
changed; the disabled path is what the other editors use for an open action.

## 3. Files

| File                                                              | Change                                              |
| ----------------------------------------------------------------- | --------------------------------------------------- |
| packages/ui/src/mods/slot-view.tsx                                | optional actions, `canRunAction`, disabled controls |
| apps/{slides,pdf,markdown,html}/src/shared/ipc.ts                 | `featuresList` on the app's API                     |
| apps/{slides,pdf,markdown,html}/src/preload/index.ts              | the one-line bridge                                 |
| apps/{slides,pdf,markdown,html}/src/renderer/mods/EditorSlots.tsx | the binding, per app                                |
| apps/{slides,pdf,markdown,html}/src/renderer/App.tsx              | the rail in the dock and the preset hand-off        |

Each app's binding points at its own global (`window.slidesApi`, `window.pdfApi`,
`window.markdownApi`, `window.htmlApi`) — the renderer never reaches across apps.

## 4. Verification

- all five editor packages typecheck with **0 errors**, the shell's own sources
  included;
- agent-core (210), shell (30) and the ui package typecheck are unchanged and green;
- theme, english-only and prettier gates pass;
- the renderer itself is already covered by the picture harness, which draws
  `editor.side` through the same component these five bindings use.

## 5. Out of scope

- **sheets**: its assistant is a different component (`AiChatPanel`) with its own
  surface, so it needs its own pass rather than a copy of this one;
- a file-opener per editor: the four bridges have none, so open actions draw
  disabled until each app exposes one;
- editor overlays, editor-side native views, and the ribbon/menu/status places.

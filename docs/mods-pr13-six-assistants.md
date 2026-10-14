# PR 13 — every editor's assistant can use the mod features

Status: **implemented**. Follows docs/mods-pr12-mods-tools.md, which did Docs first.

## 1. What ships

The same read-only mod tools in the other five editors' assistants (sheets, slides,
pdf, markdown, html). Each got four small pieces:

| Piece    | What it is                                                                                                               |
| -------- | ------------------------------------------------------------------------------------------------------------------------ |
| bridge   | `featuresRun` on the app's own API type and preload — read-only, and no apply                                            |
| snapshot | `currentEditorFeatures()` from the app's mods module: the registry as of this turn, for the assistant's per-turn context |
| skill    | `createModsSkill({ snapshot, list, dryRun })` composed beside the app's own skills                                       |
| context  | the flat "id \| name \| kind" list, labelled as data rather than instruction                                             |

Pages that compose skills with `composeSkills(...)` took one more entry; pdf, whose
assistant had a single skill, now composes its own skill plus the mods skill —
`composeSkills('pdf+mods', '', [createPdfSkill(deps), createModsSkill(...)])`.

## 2. What none of them got

An apply tool. Every bridge carries the read-only run and nothing else, so a model
in any of the six editors can inspect and dry-run the user's features and cannot
write with them. Writing is still the user, in the feature panel, after reading a
plan — the same two-step gate that shipped in docs/mods-pr11-script-run.md.

## 3. Files

| File                                                                                 | Change                         |
| ------------------------------------------------------------------------------------ | ------------------------------ |
| apps/{sheets,slides,pdf,markdown,html}/src/shared/*.ts                               | `featuresRun` on the app's API |
| apps/{sheets,slides,pdf,markdown,html}/src/preload/index.ts                          | the one-line read-only bridge  |
| apps/{sheets,slides,pdf,markdown,html}/src/renderer/mods/EditorSlots.tsx             | `currentEditorFeatures()`      |
| apps/{sheets,slides,pdf,markdown,html}/src/renderer/**/AiPanel.tsx (sheets: App.tsx) | the skill composed             |

## 4. Verification

- all six editors typecheck with **0 errors** (docs included, unchanged this round);
- the shared skill's own six tests (agent-core) cover the contract every app now
  relies on: the two tools, the plan reporting, the refusal paths, the data-not-
  instruction context, and the claim guard;
- the full build (`npm run build:all`) succeeds, which compiles every changed
  preload and renderer;
- agent-core 216, shell 120, theme / english-only / prettier gates.

## 5. Out of scope

- an apply tool for any model: deliberately absent everywhere;
- ribbon, menu and status places: they need each app's ribbon metadata.

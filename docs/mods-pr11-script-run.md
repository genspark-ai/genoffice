# PR 11 — running a script feature, and closing the loop

Status: **implemented**. This is the piece that makes the script kind usable from
the app: a mod assistant adds a script, and the user runs it against a document.

## 1. What ships

The mod manager's panel grew a **Scripts** section. For every feature whose kind is
script it shows the feature, a picker of recent documents, and two buttons:

    [ Dry run ]   [ Apply ]

- **Dry run** is read-only: the panel calls features:run and shows the plan the
  engine produced (how many operations would apply, one line each).
- **Apply** is **disabled until a dry run of that same feature against that same
  document has succeeded** — the two-step gate is in the interface, and the main
  process checks its own two gates again underneath (an explicit confirm, and that
  no editor has the file open).

The result block is tinted by outcome: a plan in the surface tone, a write in the
success tone.

## 2. Why the panel and not the assistant

The design splits the two assistants: the **mod** assistant adds features and never
touches a document; the **native** assistant (the one that owns a document and its
save path) applies. A _user_ pressing Apply is neither model writing — it is the
user's own decision, taken after reading the plan, against a file they picked. The
mod assistant's tool surface still has no write path at all.

## 3. The three keys

Three strings, in all 21 locale blocks the shell's dictionary has:

| key         | en      |
| ----------- | ------- |
| modsScripts | Scripts |
| modsDryRun  | Dry run |
| modsApply   | Apply   |

The document picker reuses the existing \`open\` key, and the plan, the failures and
the counts are rendered from the host's own answers rather than from prose.

## 4. Files

| File                                                | Change                                                            |
| --------------------------------------------------- | ----------------------------------------------------------------- |
| apps/shell/src/renderer/src/ModsAssistant.tsx       | the Scripts section, its per-feature state, and the two-step gate |
| apps/shell/src/renderer/src/mods-assistant.css      | the section's look                                                |
| apps/shell/src/renderer/src/strings.ts              | the three keys, every locale                                      |
| apps/shell/src/shared/feature-channels.ts           | \`ScriptApplyResult\` moves here, beside the channels             |
| apps/shell/src/shared/home-api.ts, preload/index.ts | \`featuresApply\` on the shell bridge                             |
| apps/shell/src/main/script-host.ts, features-ipc.ts | the type now comes from the shared module                         |
| e2e/panel-preview                                   | the harness draws a script feature and drives the flow for real   |

## 5. Evidence

The picture harness renders the real panel against a stub bridge and **clicks
through the flow** like a user: it picks a recent document, presses Dry run
(captured: the plan shows 3 / 3 with the engine's own lines, and Apply has just
become available), then presses Apply (captured: the same block in the success
tone). The bridge in the harness answers exactly what the host answers.

## 6. Out of scope

- the native assistant's own tool for this (the next step): the model plans, the
  user presses;
- a diff preview beyond the plan lines;
- approval memory: every write asks again.

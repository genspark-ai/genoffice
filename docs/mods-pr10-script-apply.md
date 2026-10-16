# PR 10 — the script kind, second slice: applying

Status: **implemented**. Follows docs/mods-pr4-script-kind.md (read-only dry run).

## 1. What ships

The host can now apply a script feature to a document:

    applyScriptToDocument({ scriptId, source, documentPath, confirm: true })

The script **still** only proposes. The host opens the document, runs the sandbox,
and only if the engine accepted the whole batch does it save with the engine's own
writer and replace the file. There is no code path by which a script writes.

## 2. The gates

| Gate              | Behaviour                                                                                                                                                          |
| ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| confirmation      | \`confirm\` must be literally true; anything else refuses before a document is opened                                                                              |
| open in an editor | the caller's \`isDocumentOpen\` check refuses a file a tab has open: writing it here would be lost on that editor's next save (wired from the shell's tab manager) |
| atomic batch      | the engine refuses a bad operation outright, so a half-applied document is never written                                                                           |
| nothing proposed  | a script that proposes nothing is an error, not a silent no-op                                                                                                     |
| one-step replace  | the bytes are written to a temporary file and renamed over the original: a reader never sees a half-written file                                                   |

## 3. What the tests prove

On a scratch copy of a real fixture (the repo's fixture is never touched):

1. without a confirmation the call refuses and the file is **byte-identical**;
2. with one, the file changes, the result reports the applied operation, and
   **reading the file back with the same reader shows the new text on disk** — the
   write is verified, not assumed;
3. an operation the engine does not know makes the call throw and leaves the file
   byte-identical;
4. a script that proposes nothing is refused, and leaves no temporary file behind.

## 4. Where this is called from

The channel (\`features:apply\`) is registered with the same validation as the rest
of the registry, and **no preload exposes it yet** — deliberately. The design gives
the _native_ assistant (the one that owns a document and its save path) the writing
role, and the mod assistant none: the feature registry's own tool surface still has
no way to write, which is the red line this whole design is built around. The next
PR wires the native assistant in the editors to this channel.

## 5. Files

| File                                      | Change                                                                      |
| ----------------------------------------- | --------------------------------------------------------------------------- |
| packages/cli/src/headless-docs.ts         | \`saveDocument\` joins the headless path                                    |
| apps/shell/src/main/script-host.ts        | \`applyScriptToDocument\`, and the snapshot/plan helpers both runners share |
| apps/shell/src/main/features-ipc.ts       | the apply handler and its two gates                                         |
| apps/shell/src/main/index.ts              | the open-document check, wired to the tab manager                           |
| apps/shell/src/shared/feature-channels.ts | the channel                                                                 |
| apps/shell/tests/script-host.test.ts      | four apply tests on a scratch copy                                          |

## 6. Out of scope

- the native assistant's tool that calls this (next PR);
- a diff/preview before writing beyond the dry run's plan;
- per-script approval memory: every apply asks again.

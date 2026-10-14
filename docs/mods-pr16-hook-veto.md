# PR 16 — the hook kind (K2), second slice: veto

Status: **implemented**. Follows docs/mods-pr14-hook-kind.md and
docs/mods-pr15-hook-activity.md.

## 1. What ships

One more event, and it is the interesting one: **document.will-save**. A hook for it
may call `$.event.veto("reason")`, and the host then refuses that save and shows the
reason.

    create_feature({
      name: "No untitled reports",
      description: "stop me saving a report with no title",
      hook: { on: "document.will-save",
               script: { source: "if (input.event.path.indexOf('untitled') !== -1) $.event.veto('the report has no title')" } }
    })

## 2. Where the "before save" point comes from

This is the part worth reviewing: the apps save through their own IPC channels
(`docs:save`, `pdf:save`, `markdown:save-image`, …), and the shell main is the
process that serves them. So the guard wraps **ipcMain.handle once, early**, before
any app registers a handler, and inspects only save-ish channel names
(`(^|:)(save|write)(-|$)`). Nothing in any app changed, and no event bus was added.

Three properties keep it safe:

- with no `document.will-save` hook, the wrapper **calls straight through** — the
  common case costs one function call;
- a hook that throws, times out or is refused is ignored: **a broken hook must never
  break a save**;
- a veto becomes a **rejection of the invoke**, which is exactly what each app
  already renders as a failed save, with the hook's reason in the message. The
  writer is never reached — the tests assert that the original handler did not run.

## 3. What a hook still cannot do

It cannot write, and it cannot veto anything else: every other event is a
notification, and `$.event.veto` on those is recorded in the activity and ignored.
The veto is a permission question the host asks — not a capability the hook holds.

## 4. Files

| File                                          | Change                                                                       |
| --------------------------------------------- | ---------------------------------------------------------------------------- |
| packages/agent-core/src/feature-hooks.ts      | `document.will-save` joins the events, and the contract says what a veto is  |
| packages/agent-core/src/features-skill.ts     | the prompt tells the assistant about the veto, and to use it sparingly       |
| apps/shell/src/main/script-host.ts            | the sandbox gains `$.event` (`name`, `path`, `veto`) and reports the vetoes  |
| apps/shell/src/main/feature-hooks.ts          | `installSaveGuard`, the channel rule, the target reader, and the live runner |
| apps/shell/src/main/index.ts                  | the guard is installed before any app handler                                |
| apps/shell/src/main/features-ipc.ts           | hands the guard the live runner                                              |
| apps/shell/src/shared/feature-channels.ts     | a record entry may carry `vetoes`                                            |
| apps/shell/src/renderer/src/ModsAssistant.tsx | the activity shows a veto with a ⊘ and the reason                            |
| apps/shell/tests/*                            | the sandbox veto, and six cases for the guard                                |

## 5. Tests

1. the sandbox hands a hook its event and collects the veto it asked for, and
   collects nothing when none was asked;
2. save-ish channels are recognised and `app:get-auto-save-default` is not;
3. the document a save is about is read from the payload, and never invented;
4. with no hook, a save calls straight through;
5. a vetoing hook stops the save with its reason, and the writer is not reached;
6. a broken hook does not stop a save;
7. a channel that is not a save is not inspected.

## 6. Out of scope

- vetoing anything but a save;
- letting a hook _change_ what is saved: it can only stop it;
- a per-hook opt-in for the guard: with no hook there is nothing to guard.

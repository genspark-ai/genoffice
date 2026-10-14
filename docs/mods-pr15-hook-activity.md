# PR 15 — the hooks' activity, in the panel

Status: **implemented**. Follows docs/mods-pr14-hook-kind.md, which made the hook
kind run but left its record unread.

## 1. What ships

The feature panel grew a **Hooks** section: the last few things the user's hooks
saw, newest first.

    HOOKS
    ┌──────────────────────────────────────────┐
    │ note-opened-reports      document.opened │
    │ weekly-report.docx · 4:37:25 AM          │
    │ opened weekly-report.docx                │
    │ blocks=24                                │
    ├──────────────────────────────────────────┤
    │ note-opened-reports      document.closed │
    │ kickoff.pptx · 4:27:55 AM                │
    │ the sandbox refused this run             │
    └──────────────────────────────────────────┘

Each entry names the hook, the event, the file (its own name, not the full path,
plus the local time), the lines the hook logged in monospace, and — in the danger
tone — the reason when a hook threw, timed out or was refused. The section is absent
until something has happened, so an untouched install shows no empty furniture.

It is a **record, not a control**: there is nothing to press, because a hook cannot
change anything. The panel refreshes it with the same rhythm as the feature list
(on mount, before a turn, and after a tool runs).

## 2. One key

| key       | en    |
| --------- | ----- |
| modsHooks | Hooks |

In all 21 locale blocks. The entries themselves are the host's own data: a hook id,
an event name, a file name, a clock and the hook's log lines.

## 3. Files

| File                                                | Change                                                              |
| --------------------------------------------------- | ------------------------------------------------------------------- |
| apps/shell/src/shared/feature-channels.ts           | the activity record's shape, beside the channel                     |
| apps/shell/src/shared/home-api.ts, preload/index.ts | `featuresHookActivity()`                                            |
| apps/shell/src/main/feature-hooks.ts                | the record's type now comes from the shared module (one definition) |
| apps/shell/src/renderer/src/ModsAssistant.tsx       | the Hooks section and its refresh                                   |
| apps/shell/src/renderer/src/mods-assistant.css      | the section's look                                                  |
| apps/shell/src/renderer/src/strings.ts              | the key, every locale                                               |
| e2e/panel-preview                                   | the harness answers with two recorded entries                       |

## 4. Evidence

The picture harness renders the real panel against a stub bridge that answers
exactly what the host answers, so the capture shows the section as the product draws
it: two entries, one with log lines and one with a failure reason.

## 5. Out of scope

- clearing the record from the panel: it is bounded at 50 entries and cleared when
  the shell disposes the registry;
- **veto**: a hook that can stop an action. That needs the host to own the action
  (a save) first, and it is the next slice;
- more events: each one needs a real emitter behind it.

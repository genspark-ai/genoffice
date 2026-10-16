# PR 14 — the hook kind (K2), first slice: observe

Status: **implemented**. This is row 5 of the order in docs/mods-design.md
section 9.2, and the first kind that lets a feature take part in what the app does
rather than only adding to its interface.

## 1. What ships

A feature of kind **hook** observes an event the app raises:

    create_feature({
      name: "Note when a report opens",
      description: "keep a line for each report I open",
      hook: { on: "document.opened", script: { source: "\$.log(input.event.path)" } }
    })

The events in this slice are the ones the host itself can name because the shell
already tracks them: **document.opened** and **document.closed**, emitted from the
one place the open-document set is published (the shell's tabs plus detached
windows). A hook is asked for by diffing that set, so nothing had to grow an event
bus and no editor had to change.

## 2. What a hook can and cannot do

It runs in **the same sandbox as a script**: plain statements, no modules, a
16 KiB source, a wall-clock budget, and no `process`, `require`, timers or
network. It is handed the event as data (`input.event = { name, path }`) and has
no document to read — an event is not a plan.

It is **read-only by construction**, and that is not a policy it could break:
the runner has no writer to call. What it says goes to its log, and the log goes
to a bounded in-memory record (the last 50 entries) that the panel can read over
`features:hook-activity`. Vetoing an action — the more interesting half of K2 —
needs the host to own that action first, so it is deliberately not here.

## 3. Files

| File                                                  | Change                                                                                                |
| ----------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| packages/agent-core/src/feature-hooks.ts              | new: the events, the hook payload, its validation (the script contract), and which hooks observe what |
| packages/agent-core/src/features-skill.ts             | the hook on create/update, the tool schema, and the prompt that explains it                           |
| packages/agent-core/tests/feature-hooks.test.ts       | four cases for the contract                                                                           |
| apps/shell/src/main/feature-hooks.ts                  | new: the runner, the bounded activity record, and the sandbox-backed run                              |
| apps/shell/src/main/feature-store.ts, features-ipc.ts | persist and validate it, register the activity channel, expose `emitHook`                             |
| apps/shell/src/main/index.ts                          | emit on the open/close diff where documents are published                                             |
| apps/shell/src/shared/feature-channels.ts             | the activity channel                                                                                  |
| apps/shell/tests/feature-hooks.test.ts                | four cases for the host                                                                               |

## 4. Verification

- the contract: the events the host raises are accepted, an invented one is
  refused by name, the script is validated by the script contract and reported as
  `hook.script`, and the matching is by event;
- the host: only the hooks for that event run, in registry order; a hook that
  throws is recorded as an error and the others still run; an event nobody observes
  does nothing; the record is bounded and clears on dispose;
- the store still refuses a malformed payload, and the IPC boundary validates it
  again (the same three-layer shape as the slot, appearance and script kinds);
- the full build succeeds, and every package typechecks.

## 5. Out of scope

- **veto**: a hook that can stop a save. That needs the host to own the save path
  first, which is the next slice;
- more events (a save finishing, a selection changing): each one needs a real
  emitter, and an event nobody raises is not offered — the same rule that keeps
  `ribbon.button` out of the place list;
- the panel's activity view: the record is readable over IPC and the hooks show in
  the registry as kind `hook`; a dedicated view is the next PR.

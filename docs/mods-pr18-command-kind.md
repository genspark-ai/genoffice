# PR 18 — the command kind (K6)

Status: **implemented**. The last of the eight kinds the design lists.

## 1. What ships

A feature of kind **command** registers an action the user invokes from a list:

    create_feature({
      name: "Tidy headings",
      description: "unfold every wrapped heading",
      command: { label: "Tidy headings",
                 description: "unfold every wrapped heading in this report",
                 action: { kind: "ask-assistant", prompt: "Tidy the headings in this report" } }
    })

**Cmd/Ctrl+K** opens the palette: type to filter, arrow to move, Enter to run,
Escape to close. A command runs the _same_ declared actions a drawn control may run
(`open`, `ask-assistant`, `panel`, `run-feature`), so this kind adds an entry point
and nothing else: where a press starts is the whole difference.

A command this host cannot serve is listed **unavailable** rather than left to fail
when pressed — the same rule the slot renderer follows for an action it cannot run.

## 2. The pieces

| File                                                  | Change                                                                                                                                |
| ----------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------- |
| packages/agent-core/src/feature-commands.ts           | new: the command contract, its validation, and `matchesCommand` (the palette's filter, unit-tested rather than buried in a component) |
| packages/agent-core/src/feature-slots.ts              | `validateSlotAction` is exported, so a command is judged by the same action rules a slot is                                           |
| packages/agent-core/src/features-skill.ts             | the command on create/update, the schema, and the prompt                                                                              |
| apps/shell/src/main/feature-store.ts, features-ipc.ts | persist and validate it (three layers, as with every other kind)                                                                      |
| apps/shell/src/renderer/src/CommandPalette.tsx        | new: the palette                                                                                                                      |
| apps/shell/src/renderer/src/Home.tsx                  | Cmd/Ctrl+K, and the palette over the shell                                                                                            |
| packages/ui/src/mods/slot-view.tsx                    | `runSlotAction` and `canRunAction` are exported, so the palette runs and judges actions exactly as a drawn control does               |
| apps/shell/src/renderer/src/strings.ts                | three keys, all 21 locale blocks                                                                                                      |
| e2e/panel-preview                                     | a palette page, rendered with the real component                                                                                      |

## 3. What is deliberately not here

- **key bindings**: a mod cannot claim a shortcut. A palette entry is discoverable
  and collision-free; a global binding is neither, and registering one would need
  conflict rules the rest of this design does not have;
- commands from the shell itself: this palette lists the commands **features**
  registered, and says so when it is empty;
- a command that writes: the declared actions still cannot, on purpose.

## 4. Tests

1. the contract accepts all four declared actions and refuses an invented one;
2. it refuses a missing label, an over-long label, an unknown field (a key binding)
   and a malformed action, naming the field each time;
3. the filter matches over the label and the description, case-insensitively, and an
   empty query matches everything;
4. the registry lists only the features that have a command.

## 5. Evidence

The picture harness renders the real palette with real commands and captures it:
the scrim, the query field, three commands with their descriptions, and the first
one active.

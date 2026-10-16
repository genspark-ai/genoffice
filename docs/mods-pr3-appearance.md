# PR 3 — the appearance kind (K8)

Status: **implemented**. Follows docs/mods-pr2-declarative-slots.md.

## 1. What ships

A feature of kind **appearance** asks the shell's own chrome to change:

    create_feature({
      name: "Night shift",
      description: "softer chrome after dark",
      appearance: { background: "surface-subtle", accent: "violet" }
    })

The shell applies the newest appearance in the registry; removing that feature
hands its defaults back. Nothing else about the app changes.

## 2. The contract

| Field      | Values                                        |
| ---------- | --------------------------------------------- |
| background | default, surface, surface-subtle, accent-tint |
| accent     | default, blue, violet, teal, amber, rose      |

**Names only, never values.** The literals live in the shell's stylesheet, so dark
mode, contrast and the theme-token rules stay the shell's business; a raw colour
such as #ff00ff is refused with the field named. The same validation runs in the
tool, the store and the IPC boundary, so an invalid choice cannot be written or
drawn.

## 3. What it may and may not touch

It re-points **chrome variables only**: the background behind the content areas
and the accent family (buttons, chips, focus rings, selection tints).

It cannot reach: a document surface, a page background, a highlight colour map, a
chart palette, a stamp or WordArt preset, or the print and export paths. Those are
document data (see the theming rules in CLAUDE.md), and a mod never sees them.

## 4. Files

| File                                                | Change                                                                   |
| --------------------------------------------------- | ------------------------------------------------------------------------ |
| packages/agent-core/src/feature-appearance.ts       | new: the contract, validation, and which appearance is active            |
| packages/agent-core/src/features-skill.ts           | the appearance payload on create/update, and the prompt that explains it |
| apps/shell/src/main/feature-store.ts                | persist it, validate it on write and on read                             |
| apps/shell/src/main/features-ipc.ts                 | validate it at the renderer boundary                                     |
| apps/shell/src/renderer/src/slots/features-state.ts | apply the active appearance to the document root                         |
| apps/shell/src/renderer/src/slots/slots.css         | the host's palette and background definitions                            |

## 5. Tests

1. the contract accepts every name, refuses a raw colour, an unknown field and an
   empty appearance, and picks the most recently changed feature;
2. the skill records one, says plainly that the shell applies it, and refuses a
   raw colour by field;
3. the store round-trips one, gives it back on an explicit null, and skips a
   hand-written malformed one;
4. the flow: a scripted model registers one and the registry reports it as the
   active appearance — which is exactly what the renderer reads.

## 6. Out of scope

- per-app accents: this is the shell's chrome;
- background images, gradients or any value the user supplies;
- theming a document surface (the dark-page feature already does that as a
  display-only remap, and it is not a mod).

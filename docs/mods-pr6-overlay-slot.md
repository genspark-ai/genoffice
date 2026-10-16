# PR 6 — the canvas overlay slot (K4)

Status: **implemented**. The third of the three slots: the declarative one shipped
in docs/mods-pr2-declarative-slots.md, the native view in docs/mods-pr5-native-view.md.

## 1. What ships

A screen-only layer over the content area:

    create_feature({
      name: "Review guide",
      description: "Points at the parts of a report that need a reader.",
      slot: { at: "home.overlay", element: { type: "overlay", shapes: [
        { kind: "frame", x: 0.5, y: 0.08, width: 0.47, height: 0.24, tone: "success", dashed: true },
        { kind: "dot", x: 0.5, y: 0.62, tone: "danger", label: "A page the mod added" }
      ] } }
    })

## 2. The contract

| Field               | Rule                                                                    |
| ------------------- | ----------------------------------------------------------------------- |
| shapes              | at least one, at most 64                                                |
| kind                | frame, line, dot or label                                               |
| x, y, width, height | fractions of the layer (0 to 1), never pixels                           |
| tone                | accent, muted, success or danger — the four status tokens the shell has |
| label / value       | at most 80 characters                                                   |

Fractions are the point: a mod never measures a document, so an overlay cannot
claim a pixel position it did not derive. Tones are names, so a mod never picks a
colour and dark mode stays correct.

## 3. What the host enforces

- the layer is **inert**: no pointer events, `aria-hidden`, no focus — every press
  reaches the app underneath, so a mod can point at something but cannot add a
  control it does not own;
- the layer is **screen-only**: it is DOM chrome over the content area. A document,
  its rendering and the print/export paths never see it — there is no code path
  from here into a document;
- tones resolve to `--accent`, `--text-tertiary`, `--success` and `--danger`
  through classes in the shell's stylesheet, which is why the theme gate stays green;
- the content area only becomes a positioning context while an overlay exists
  (`content--overlay`), so with no overlay the layout is byte-for-byte what it was.

## 4. Files

| File                                               | Change                                           |
| -------------------------------------------------- | ------------------------------------------------ |
| packages/agent-core/src/slot-overlay.ts            | new: the shape contract and its validation       |
| packages/agent-core/src/feature-slots.ts           | the `home.overlay` place and the overlay element |
| apps/shell/src/renderer/src/slots/OverlayLayer.tsx | new: the layer and the shapes                    |
| apps/shell/src/renderer/src/slots/slots.css        | the frame, the guides and the tone classes       |
| apps/shell/src/renderer/src/Home.tsx, home.css     | mount the layer over the content area            |

## 5. Tests

1. the contract accepts a frame, a guide, a dot and a label;
2. it refuses a coordinate outside the layer, naming the field; an invented tone;
   an unknown shape kind; an empty layer; more than 64 shapes; a long label;
3. the slots contract accepts `home.overlay` and refuses a bad shape inside it;
4. the picture harness draws a real overlay over the gallery.

## 6. Out of scope

- interaction: hit areas, hover, tooltips, drag. An overlay that takes presses is a
  different, larger question (who owns the press, and what may it do);
- measuring the document for the mod: the honest answer is the native assistant
  measures and the mod declares fractions, which is what this contract assumes;
- overlays inside the editors: same contract, but they need the shared surface work
  from the fan-out.

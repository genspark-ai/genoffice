# PR 7 — editor slots: one renderer, two apps

Status: **implemented**. Opens the first editor place and moves the slot renderer
into @genoffice/ui so the shell and every editor draw a slot the same way.

## 1. What ships

1. **The renderer moves into the shared package.** `SlotHost`, `SlotPanel` and
   `OverlayLayer` now live in packages/ui/src/mods with their stylesheet
   (mods.css). The package owns rendering, the element set and the styles; the app
   owns the bridge. There is exactly one implementation of "what a button in a
   slot looks like".
2. **The injected surface.** A host passes `SlotActions`
   (open, ask, panel, listSource, mountView, unmountView) and `SlotLabels.close`.
   Nothing in the package reaches for `window`, so a slot cannot smuggle in a
   bridge the app did not hand it.
3. **`editor.side`**, the first editor place: a section in an editor's own dock.
4. **The docs app binds it**: the registry (read-only, one channel), a rail inside
   the AI dock, and what a press does in that editor — open a document with the
   editor's own open path, or hand the prompt to the editor's own assistant
   (its preset, which runs immediately).

## 2. Why shared rather than copied

Every app has its own preload, so a slot renderer per app would be six copies of
the same 400 lines, six stylesheets and six chances to drift. The package is
already a dependency of the shell and of every editor, so the renderer lives there
and each app contributes only what is genuinely its own: where its registry comes
from and what its bridge can do.

The shell keeps exactly one thing of its own: the appearance rules
(apps/shell/src/renderer/src/slots/appearance.css), because they re-point the
shell's chrome variables rather than render a slot.

## 3. A real fix found while moving

The primary button in a slot had no on-accent colour (`--mods-on-accent` was a
shell-local literal), so once the stylesheet was shared the button's label fell
back to the surface colour and vanished on the accent fill. There is now a proper
`--on-accent` token in packages/ui/src/tokens.css, in all three blocks (light,
dark and the system-dark fallback), and both the shared button and the shell's
composer use it.

## 4. Fidelity

The picture harness renders through the shell's binding. Its captures before and
after the move are the same picture — same frames, same overlay, same accent —
which is the evidence that moving the renderer changed no pixel. The gallery now
also draws `editor.side` through the same component, to show the editor place.

## 5. Files

| File                                                            | Change                                                                   |
| --------------------------------------------------------------- | ------------------------------------------------------------------------ |
| packages/ui/src/mods/slot-view.tsx, overlay-layer.tsx, mods.css | new: the shared renderer and its styles                                  |
| packages/ui/src/tokens.css                                      | the `--on-accent` token, three blocks                                    |
| packages/ui/src/index.ts                                        | the package's exports                                                    |
| packages/agent-core/src/feature-slots.ts                        | the `editor.side` place; `viewHost` moves here beside `isAllowedViewUrl` |
| apps/shell/src/renderer/src/slots/*                             | the shell's binding: store and bridge in, renderer out                   |
| apps/shell/src/renderer/src/slots/appearance.css                | the shell's own chrome rules (renamed from slots.css)                    |
| apps/docs/src/shared/ipc.ts, preload/index.ts                   | `featuresList` on the editor's bridge                                    |
| apps/docs/src/renderer/mods/EditorSlots.tsx                     | new: the editor's binding                                                |
| apps/docs/src/renderer/App.tsx                                  | the rail in the AI dock, and the preset hand-off                         |

## 6. Verification

- agent-core: 209 tests, including the place list (`editor.side` accepted, an
  unknown place refused with the list of real ones);
- shell: 30 tests; the shell's own typecheck is clean and its preview captures are
  unchanged;
- docs: the package's typecheck is clean (0 errors), the editor's binding included;
- theme, english-only and prettier gates pass.

## 7. Out of scope

- the other five editors: same pattern, one PR each;
- ribbon, menu and status places: they need each app's ribbon metadata, and a place
  is only added to the contract when a host actually draws it — which is why
  `panel.header` is not in the list yet;
- editor overlays and editor-side native views: the contracts already cover them,
  the hosts come with the fan-out.

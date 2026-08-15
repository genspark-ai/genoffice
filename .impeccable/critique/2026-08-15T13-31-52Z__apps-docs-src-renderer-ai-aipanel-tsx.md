---
target: apps/docs/src/renderer/ai/AiPanel.tsx
total_score: 16
max_score: 40
na_heuristics:
p0_count: 0
p1_count: 4
timestamp: 2026-08-15T13-31-52Z
slug: apps-docs-src-renderer-ai-aipanel-tsx
---

⚠️ DEGRADED: single-context (sub-agents declined by user)

# Codex selector critique

Scope: the three supplied screenshots and the Docs Codex selector/composer implementation. The Provider settings surface is not visible here, so this critique does not judge it.

## Design Health Score

| #         | Heuristic                       |     Score | Key issue                                                                                                                                                    |
| --------- | ------------------------------- | --------: | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 1         | Visibility of System Status     |       2/4 | The selected values are visible, but the open menu does not clearly communicate its spatial relationship to the trigger, and edge clipping gives no warning. |
| 2         | Match System / Real World       |       3/4 | The Model/Effort/Speed vocabulary and model names feel Codex-specific; the geometry does not follow the reference convention.                                |
| 3         | User Control and Freedom        |       2/4 | Escape, outside-click, and reset exist, but a clipped menu can make choices unreachable.                                                                     |
| 4         | Consistency and Standards       |       1/4 | The root and submenu use different proportions and a detached-card treatment instead of one attached context menu.                                           |
| 5         | Error Prevention                |       1/4 | Fixed absolute placement has no viewport collision, flip, or clamp behavior.                                                                                 |
| 6         | Recognition Rather Than Recall  |       2/4 | Labels are clear, but the offset submenu makes users infer which root row owns it.                                                                           |
| 7         | Flexibility and Efficiency      |       2/4 | Click, hover, focus, and Escape are supported, but there is no directional keyboard navigation and each selection closes the menu.                           |
| 8         | Aesthetic and Minimalist Design |       1/4 | The root menu is over-tall, the outlines are heavy, and the overlay dominates the composer/editor.                                                           |
| 9         | Error Recovery                  |       1/4 | There is no visible recovery when the menu extends outside the available area.                                                                               |
| 10        | Help and Documentation          |       1/4 | Speed and service-tier meaning are not explained at the decision point.                                                                                      |
| **Total** |                                 | **16/40** | **Poor — significant improvement required before shipping.**                                                                                                 |

## Design Specificity Verdict

The copy is authored for Codex; the composition is not yet authored for either Codex or GenOffice. It currently reads as two generic dark CSS cards placed near a composer, rather than one contextual menu whose submenu grows naturally from the hovered row.

The source makes the main causes explicit in `AiPanel.tsx` and `styles.css`: the popover is a fixed-width flex row with `right: 0`, `align-items: flex-end`, `gap: 8px`, strong borders, and fixed row heights. That combination explains all three screenshots:

- In the open state, the submenu is bottom-aligned with the root, so its top begins substantially lower than the root. It looks detached instead of like the child menu of Model.
- The 8px gap creates a visible seam between two separately rounded cards. The reference uses a flush, joined surface.
- The root is almost as wide as the submenu, while the reference submenu is materially wider. The current proportions make the root feel inflated and the option panel cramped.
- In the edge-positioned state, the combined fixed footprint is wider than the available dock. `right: 0` pushes the root off-screen while `overflow: visible` merely allows the submenu to spill over the document; it does not solve placement.

The deterministic detector reported zero findings for the TSX target. That is a false sense of safety here: the defects are runtime geometry, CSS proportions, and viewport collision behavior, which the markup scan cannot evaluate.

## Overall Impression

The implementation is functionally recognizable, but visually broken. The strongest failure is not color or copy; it is that the menu has no coherent spatial system. A user expects one attached Codex menu and instead sees a tall card plus a lower, detached card that can clip at the dock edge.

## What's Working

1. The closed trigger is compact and expands to the root-menu width, which matches the requested interaction direction.
2. The three controls are correctly named Model, Effort, and Speed, with Codex-style model/effort/speed values.
3. The selected model, reset action, Escape handling, and outside-click dismissal provide a reasonable behavioral foundation once the geometry is corrected.

## Priority Issues

### [P1] The submenu is anchored to the bottom, not to the context

**Why it matters:** In the open screenshot the root begins around the top of the overlay while the submenu begins much lower. The user cannot immediately see that the right panel belongs to the highlighted Model row. This breaks the defining affordance of the reference interaction.

**Fix:** Make the default Model submenu share the root's top edge. For Effort and Speed, anchor the submenu to the active root row or use a deliberate row-relative offset. Do not bottom-align the two panels.

**Suggested command:** `$impeccable layout`

### [P1] The fixed two-panel footprint has no viewport collision strategy

**Why it matters:** The third screenshot shows the root menu clipped by the left edge while the submenu floats over the editor. The same failure will occur near other edges and at smaller panel widths. A control that exposes choices but hides part of its own menu is effectively broken.

**Fix:** Position the menu as a collision-aware floating layer. Clamp it to the window, flip the submenu side when needed, and keep the active root row visible. Allow document overlap only when the complete menu remains visible.

**Suggested command:** `$impeccable adapt`

### [P1] The root and submenu do not read as one surface

**Why it matters:** The visible gap, separate borders, separate radii, and nearly identical widths make the two panels look like unrelated popovers. The reference reads as one joined context menu with a clear parent/child relationship.

**Fix:** Remove the gap, share the seam, use the root width for the control panel and a meaningfully wider submenu, and reserve outer corner radii for the outside edges of the combined surface.

**Suggested command:** `$impeccable distill`

### [P1] The root menu's vertical rhythm is inflated and inconsistent

**Why it matters:** The root has three large rows plus a large reset row, while the submenu uses a different density. The result is a tall empty card that dominates the composer and makes the selector feel like a modal rather than a quick control.

**Fix:** Define one compact menu-row rhythm and use it for root rows, reset, and submenu options. Reduce root row height/padding first; preserve enough hit area through the actual interactive box, not blank space between text.

**Suggested command:** `$impeccable typeset`

### [P2] Focus and interaction states are under-specified

**Why it matters:** Hover changes the active submenu, but there is no directional keyboard model for moving between root and child menus. The root menu's accessible label is always Model even when Effort or Speed is active. Closing after every choice also forces repeated open/close cycles to configure three settings.

**Fix:** Add explicit active-row-to-submenu semantics, arrow-key navigation, and correct accessible labels. Decide whether the menu should remain open while users move between Model, Effort, and Speed; the reference suggests a single inspection surface rather than three separate trips.

**Suggested command:** `$impeccable harden`

## Persona Red Flags

### Alex — Impatient Power User

- Must reopen the selector after choosing a model before changing Effort or Speed.
- Cannot use arrow keys to move from a root row into its submenu; focus support is partial rather than a complete menu interaction.
- The 180ms width animation is fine in isolation, but the large repositioned overlay creates unnecessary visual movement for a frequent control.

### Jordan — Confused First-Timer

- The lower-aligned submenu does not visually explain which of Model/Effort/Speed it belongs to.
- The detached panels and large empty root make the UI look broken or like two competing dialogs.
- At the dock edge, clipping removes the beginning of the menu with no explanation or recovery cue.

### Sam — Accessibility-Dependent User

- Hover is the strongest visible interaction cue, but keyboard users do not get equivalent arrow-key traversal or a linked `aria-controls` relationship.
- The selected blue check and neutral selected background are not enough to explain the full menu hierarchy when the submenu is spatially displaced.
- The root menu label remains Model while Effort or Speed is active, which can produce misleading screen-reader output.

## Minor Observations

- The bright blue check is much louder than the neutral Codex reference; a quieter semantic accent would better match the monochrome menu.
- Strong borders around both panels compete with the existing dark Docs surfaces; a subtler border and more controlled elevation would feel more native to GenOffice.
- The side panel's width should be visibly larger than the root panel; model names are the long-content case and deserve that space.
- The current screenshots show the selector overlapping the composer/editor boundary. That overlap can be intentional, but only after collision handling guarantees that the send control and status bar remain legible.

## Questions to Consider

- Should this be implemented as one joined surface with a shared seam, rather than two independently rounded cards?
- When the selector is near the AI dock edge, should it flip/clamp while still overlaying the document, or must the entire menu remain inside the dock?
- Should selecting Model/Effort/Speed leave the menu open for rapid multi-setting configuration, or should it continue closing after every selection?

# Route-B prototype research: per-page rendering for large documents

**Branch**: `feat/layout-parking-proto` (local only — not pushed, no PR, per the user's
instruction). **Issue**: #2781 (large documents laggy). **Status**: working prototype,
measured A/B in real Chromium; the verdict is shape-dependent (see section 6).

This note is the handoff for the architectural decision. It maps what the current
renderer does, what the prototype changes, what stays broken, and what a maintainer
would need to decide before this becomes a real PR.

## 1. Where the continuous-DOM cost actually lives

The renderer keeps the whole document as one continuous ProseMirror DOM. Pagination is
display-only: zero-height `.page-gap-cut` markers and gap widgets at measured offsets.
Everything about "pages" is derived, never structural. Consequences, by cost center
(measured on a synthetic 3000-paragraph / ~100-page / 6000-node document, Apple
Silicon, jsdom for the walks — real numbers in `#2793`):

| cost                                                           | scales with | measured                                                           | note                                                                                   |
| -------------------------------------------------------------- | ----------- | ------------------------------------------------------------------ | -------------------------------------------------------------------------------------- |
| per-keystroke JS walks (list backfill, CJK factor, word count) | O(document) | 0.67 ms / 1.1 ms / 11.4 ms                                         | fixed in #2793 (touched-range gates)                                                   |
| Chromium style recalc + layout + paint of the whole tree       | O(document) | not measurable in jsdom                                            | the true O(document) term; broad descendant selectors re-match across the tree         |
| the whole-document measurement pass                            | O(document) | hundreds of ms bursts (see `__pageDebug.remeasureMs` in the probe) | already adaptively scheduled (300 ms – 2 s duty); Word paginates in the background too |
| ProseMirror view update (DOM diff)                             | O(changed)  | small                                                              | incremental already                                                                    |
| React re-render of App.tsx per keystroke                       | O(App tree) | doc-size independent                                               | not this issue                                                                         |

The measurement pass is engineered hard (FollowUpChain, resume-from-dirty-page, fast
passes for small edits, `measuring-natural` sampling state, layout health checks). It
is not the thing to rewrite first. The unfixable-in-route-A term is Chromium's
tree-scale style recalc/layout/paint, which grows with total DOM — that is what route
B targets.

## 2. What the prototype does (implemented here)

**Layout parking**: between measurement passes, offscreen top-level blocks carry
`content-visibility: auto` + `contain-intrinsic-size: auto <remembered height>`.
Chromium skips parked subtrees entirely (style, layout, paint) — the typing cost stops
growing with the document. The caret entering a parked block reveals it (spec:
content is skipped until "relevant to the user" — on-screen, focused, selected).

- `apps/docs/src/renderer/editor/layout-parking.ts` — pure decisions + DOM applier:
  `parkDecisions(boxes, band)`, `parkBand(scroller)` (viewport grown by one viewport
  of headroom), `applyParking`, `suppressParking` / `resumeParking`.
- CSS: `.ProseMirror .cv-parked` and the `.parking-suppressed` override, next to the
  existing `measuring-natural` rules.
- Wiring in the pagination effect (`App.tsx`): every pass suppresses parking before
  sampling (fast passes included) and re-parks in its `finally` (the pass reads real
  boxes; afterwards parked boxes report remembered sizes). A rAF-coalesced scroll
  listener shifts the band. Guards: parking is active only for the continuous
  single-column print canvas (`!hasVertical && colMode === 'none' && viewMode ===
'print' && !readMode`); `__layoutParkingOff` (window flag) disables it — the e2e
  probe uses this for the A/B run. Export/print-preview is already excluded by the
  pass deferral under `.pv-exporting`.
- Tests: `tests/layout-parking.test.ts` (decisions exact from both sides, remembered
  height stamping, band headroom, suppression lifecycle). Regression subset green
  (123 tests incl. pagination-gaps / story-view).

**Why `content-visibility` and not true virtualization**: ProseMirror requires the
full document in the DOM; unmounting offscreen content breaks the editor model. True
per-page DOM fragmentation (a page node tree) would be a different editor — the honest
first step on this route is Chromium's own skip mechanism, which keeps the document
model and the PM view intact. Real editors (CKEditor 5 long-document mode et al.) ship
exactly this CSS mechanism for the same problem.

## 3. What the probe measures (e2e/large-doc-parking.spec.ts)

Opens a generated ~100-page docx twice (parking off / on, A/B by the window flag):

- `parked` count after the open settles (mechanism: > 80% of blocks offscreen);
- caret into the middle of the document (a parked region): computed
  `content-visibility` must not be `auto` while the caret is inside (reveal works),
  typed text must land;
- scroll to the far end: content reveals on arrival;
- typing jank: 400 requestAnimationFrame gaps sampled while ~72 characters land —
  median / p95 / max / count of >100 ms frames, printed for both runs;
- `__pageDebug.remeasureMs` after the open settle and after typing (the pass burst
  the adaptive scheduler amortizes).

The spec pins the mechanism and prints the numbers; it is local-only (a latency
assertion would be machine-dependent flake in CI).

## 4. Known prototype gaps (deliberate, listed for the decision)

1. **Zoom**: `--park-h` stamps the unzoomed measured height; the pass already swaps
   zoom to 1 while sampling, but a zoomed canvas between passes would need the stamp
   divided by the live zoom factor. Prototype targets zoom 1; the zoom-swap machinery
   is the natural place to generalize.
2. **Drift between passes**: a parked block's box reports the remembered height; an
   edit that changes a block's height shifts later boxes by the delta until the next
   pass re-parks. `contain-intrinsic-size: auto` makes Chromium remember last real
   size, so drift is usually zero; worst case the scrollbar shifts on the pass after
   an edit — same behavior class as Word's background repagination.
3. **Find-in-page / comment balloons / AI queue anchors**: any UI that scrolls to an
   offscreen position relies on Chromium's reveal-on-scroll (spec'd for skipped
   content) — verified for plain scroll in the probe; balloon anchoring against parked
   blocks needs a dedicated pass before this leaves prototype status.
4. **Split view**: the read-only copy is a separate render and stays unparked; it
   could park identically later.
5. **`measuring-natural` interplay**: parking suppression wraps the whole pass
   including the `measuring-natural` sampling state — the two toggles are
   orthogonal and the finally re-parks after `locate()`; a pass that throws still
   resumes parking (finally).

## 5. The decision this research feeds (#2651 family)

Route A (shipped) keeps the continuous DOM; route B in its full form (per-page
fragment rendering, page nodes in the schema, DOM virtualization of offscreen pages)
would let layout/paint cost drop from O(document) to O(viewport) — but it rewrites the
pagination/canvas layer everything sits on (page gaps, headers/footers, print preview,
caret marks, split view). This prototype demonstrates the cheap 80%: Chromium's
content-visibility gives the O(viewport) typing/paint win _without_ any structural
change, leaving the true fragmentation question (row fragments across page breaks —
#2651's actual ask) decoupled from the performance question (#2781).

If the probe shows the parked run's typing frame gaps near the small-document
baseline, route B's performance motivation is satisfied by parking alone, and #2651's
row-fragmentation can be judged purely on its fidelity merits. If the probe still
shows long frames, the residual is the measurement pass + PM view update, and a real
fragmentation discussion (with its rewrite cost) becomes the only remaining lever.

## 6. Measured results (2026-11-17, Apple Silicon, real Chromium via the probe)

**Pure text, 5000 paragraphs / 183 pages / 5182 blocks** (`/tmp/probe-ab.js`, A/B):

|                               | parking OFF  | parking ON          |
| ----------------------------- | ------------ | ------------------- |
| parked after open             | 0            | 4957 (95.6%)        |
| typing RTT median / p95 / max | 1 / 1 / 4 ms | 1 / **94** / 107 ms |
| whole-document pass           | 166 ms       | **441 ms**          |

**Table-heavy, 40 tables x 30 bordered rows x 6 cells (~25 pages) / 122 blocks**:

|                               | parking OFF    | parking ON         |
| ----------------------------- | -------------- | ------------------ |
| parked after open             | 0              | 118 (96.7%)        |
| typing RTT median / p95 / max | 4 / 21 / 41 ms | **2 / 13 / 16 ms** |
| whole-document pass           | 334 / 309 ms   | **295 / 151 ms**   |

Readings, honest:

1. **A plain-text 100+ page document does not lag** on this machine — the
   per-keystroke main-thread RTT is 1 ms with parking off. The #2781 reporter's
   lag must come from a heavier shape (tables/images/complex styles), an older
   version, or Windows. The parked run's p95=94 ms is the **pass wake-up tax**:
   resuming a parked tree costs a full relayout of it, and a typing-triggered
   pass burst (441 ms) lands inside the sampled window. On cheap-to-layout
   documents parking is a net negative.
2. **On layout-heavy documents parking wins on every axis** — typing RTT AND
   the measurement pass itself (the pass also skips parked subtrees while
   sampling-suppressed is off... precisely: the suppression lifts only around
   sampling, so between passes the tree stays small). The heavier each block's
   layout, the bigger the win — which is exactly the #2781 shape (tables and
   pictures across 100+ pages).
3. The caret-reveal path works: the caret typed into a parked region lands and
   renders (`typedOk` true in both runs); Chromium reveals focused content per
   the content-visibility spec.

Design consequence: parking must engage **by block layout cost**, not document
length — implemented here as the sticky per-block cost gate
(`parkingArmDecided`, layout-parking.ts): arm when the last full pass cost

> = 0.25 ms/block, sticky (parked passes are cheaper, a plain re-check would
> flap), disarm below 50 blocks. **Verified with the gate in the build**: the
> plain-text 183-page document parks nothing (parked = 0, typing RTT p95 = 3 ms,
> no wake-up tax) and the table document parks 118/122 (typing RTT p95 = 9 ms,
> post-typing pass 155 ms). Calibration constants are knobs, not physics; a
> mixed text-and-tables shape dilutes the per-block average and may need a
> per-block-cost refinement later.

## 7. How to verify locally

```bash
# unit tests
cd apps/docs && npx vitest run tests/layout-parking.test.ts
# build, then the probe (A/B printed to stdout)
npm run build:all
npx playwright test e2e/large-doc-parking.spec.ts
```

Toggle parking off at runtime for manual comparison: in the docs window console
`window.__layoutParkingOff = true`, then edit anything to trigger a pass.

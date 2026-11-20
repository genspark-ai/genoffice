import { Extension } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Plugin, PluginKey } from '@tiptap/pm/state'
import type { EditorView } from '@tiptap/pm/view'
import { Decoration, DecorationSet } from '@tiptap/pm/view'

/**
 * Layout parking — route-B prototype for the large-document lag (#2781).
 *
 * Between measurement passes, offscreen top-level blocks carry
 * `content-visibility: auto` with a remembered intrinsic height, so
 * Chromium skips their style recalc, layout and paint entirely: typing
 * cost stops growing with the document. The caret entering a parked block
 * reveals it (spec: skipped content stops being skipped when it is
 * on-screen, focused or selected).
 *
 * Opt-in: this extension never parks on its own — entries arrive only via
 * the parkingKey meta, and the App side dispatches that meta only when the
 * user has enabled parking in Preferences (layout-parking-pref) and the
 * sticky per-block cost gate (parkingArmDecided) has armed.
 *
 * The parked state lives as ProseMirror **node decorations**, not as CSS
 * classes stamped from the outside: tiptap re-asserts its editor
 * attributes on every view flush and the measurement passes rebuild gap
 * widgets, so anything parked through `classList` from the App side is
 * wiped within one flush (observed: 5143 stamped blocks, zero after one
 * IPC boundary). Decorations are PM-managed state — they map across
 * transactions and are re-applied to whatever DOM PM renders.
 *
 * Cost shape, honestly: entries are mapped across every transaction (O(n)
 * position walks, n = parked blocks) and the decoration set is rebuilt
 * whenever the document or the entry list changes (O(n) `nodeAt` lookups).
 * Only selection-only transactions reuse the previous set. That is real
 * work on the typing path — acceptable because n stays in the low
 * thousands only for documents that opted in and measured expensive, and
 * because it replaces the O(document) style recalc/layout/paint the
 * prototype exists to remove.
 *
 * The measurement pass must sample real layout, so parking rides under a
 * `parking-suppressed` class on the editor element for the pass (the CSS
 * forces `content-visibility: visible`); the caller toggles it around the
 * pass and dispatches a fresh parking decision at the pass end.
 */

export interface ParkEntry {
  /** PM position of a top-level block's opening token */
  pos: number
  /** last real rendered height, px */
  h: number
}

export interface ParkingState {
  entries: ParkEntry[]
}

export const parkingKey = new PluginKey<ParkingState>('layoutParking')

/** forces parked decorations visible while a pass samples real layout */
export const PARK_SUPPRESSED_CLASS = 'parking-suppressed'

/** Parking pays a wake-up tax: resuming a parked tree costs a full relayout
 *  of it, and the first measurement pass after parking inflates by exactly
 *  that. Measured A/B (#2781 research): a plain-text 183-page document runs
 *  ~0.03 ms of pass per block with parking off — parking there is a net
 *  negative — while a table-heavy document runs ~2.4 ms/block and wins on
 *  every axis. The arm decision is therefore per-block pass cost, sticky:
 *  once a document has proven expensive to lay out it stays armed (parked
 *  passes are cheaper, so a plain re-arm check would flap), and it disarms
 *  when the document shrinks below the floor. */
export const PARK_ARM_PER_BLOCK_MS = 0.25
export const PARK_MIN_BLOCKS = 50

export function parkingArmDecided(
  lastPassMs: number,
  blockCount: number,
  previouslyArmed: boolean,
): boolean {
  if (blockCount < PARK_MIN_BLOCKS) return false
  if (previouslyArmed) return true
  return lastPassMs / Math.max(1, blockCount) >= PARK_ARM_PER_BLOCK_MS
}

export interface ParkingUpdate {
  entries: ParkEntry[]
}

/**
 * Which boxes park, from screen-space boxes and the viewport band: a block
 * parks when it sits entirely above or below the band. A box straddling the
 * edge stays unparked — the seam must be real layout, and Chromium reveals
 * skipped content lazily anyway.
 */
export function parkDecisions(
  boxes: Array<{ top: number; height: number }>,
  band: { top: number; bottom: number },
): boolean[] {
  return boxes.map((b) => b.top + b.height <= band.top || b.top >= band.bottom)
}

/** the viewport band in document coordinates: scroll position grown by one
    viewport of headroom each way, so normal scrolling lands inside
    already-real layout and only fast scrolling meets a reveal */
export function parkBand(scroller: { scrollTop: number; clientHeight: number }): {
  top: number
  bottom: number
} {
  return {
    top: scroller.scrollTop - scroller.clientHeight,
    bottom: scroller.scrollTop + scroller.clientHeight * 2,
  }
}

export const LayoutParkingExtension = Extension.create({
  name: 'layoutParking',

  addProseMirrorPlugins() {
    // the decoration set, rebuilt only when the entries or the document
    // actually change: selection-only transactions and scroll-driven view
    // updates re-read the previous set instead of walking every entry again
    let cached: {
      entries: ParkEntry[]
      doc: ProseMirrorNode
      set: DecorationSet
    } | null = null

    const decorationsFor = (doc: ProseMirrorNode, entries: ParkEntry[]): DecorationSet => {
      if (cached && cached.entries === entries && cached.doc === doc) return cached.set
      const decos: Decoration[] = []
      for (const entry of entries) {
        const node = doc.nodeAt(entry.pos)
        if (!node) continue
        decos.push(
          Decoration.node(entry.pos, entry.pos + node.nodeSize, {
            class: 'cv-parked',
            style: `--park-h:${Math.round(entry.h)}px`,
          }),
        )
      }
      const set = DecorationSet.create(doc, decos)
      cached = { entries, doc, set }
      return set
    }

    return [
      new Plugin<ParkingState>({
        key: parkingKey,
        state: {
          init: () => ({ entries: [] }),
          apply(tr, prev) {
            const update = tr.getMeta(parkingKey) as ParkingUpdate | undefined
            if (update) return { entries: update.entries }
            if (!prev.entries.length) return prev
            // map the parked positions across the transaction; an entry whose
            // position merged away is dropped and re-derives at the next pass
            const entries: ParkEntry[] = []
            let moved = false
            for (const entry of prev.entries) {
              const pos = tr.mapping.map(entry.pos, -1)
              if (pos !== entry.pos) moved = true
              if (pos >= 0 && pos <= tr.doc.content.size) entries.push({ pos, h: entry.h })
            }
            // positions untouched (selection or decoration-only transactions):
            // keep the previous array so the decoration cache stays valid
            if (!moved && entries.length === prev.entries.length) return prev
            return { entries }
          },
        },
        props: {
          decorations(state) {
            const { entries } = parkingKey.getState(state) ?? { entries: [] }
            if (entries.length === 0) return DecorationSet.empty
            return decorationsFor(state.doc, entries)
          },
        },
        view(_view: EditorView) {
          // parked blocks whose positions mapped past the document end would
          // linger forever; the pass re-dispatches anyway, this is a tidier
          return {
            update(sweepView: EditorView) {
              const state = parkingKey.getState(sweepView.state)
              if (!state || state.entries.length === 0) return
              const alive = state.entries.filter((e) => {
                const node = sweepView.state.doc.nodeAt(e.pos)
                return Boolean(node)
              })
              if (alive.length !== state.entries.length) {
                sweepView.dispatch(sweepView.state.tr.setMeta(parkingKey, { entries: alive }))
              }
            },
          }
        },
      }),
    ]
  },
})

/** helper for the App side: build the parking update from top-level blocks.
    `boxes[i]` corresponds to `doc.child(i)` (the pass measures pm.children,
    which render the PM top-level nodes one to one); the caller has already
    decided which boxes park. */
export function parkingEntriesFor(
  doc: ProseMirrorNode,
  decisions: readonly boolean[],
  heights: readonly number[],
): ParkEntry[] {
  const entries: ParkEntry[] = []
  let pos = 0
  const limit = Math.min(decisions.length, doc.childCount, heights.length)
  for (let i = 0; i < limit; i++) {
    const node = doc.child(i)
    if (decisions[i]) entries.push({ pos, h: heights[i] ?? 0 })
    pos += node.nodeSize
  }
  return entries
}

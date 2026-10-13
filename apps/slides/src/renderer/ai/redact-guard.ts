import { placeholderSource } from './redact'
import type { RenderNode, RenderSlide } from '@genoffice/pptx-render'

/**
 * Stop a model edit from damaging a span the reader withheld.
 *
 * ## What is at stake
 *
 * A mark is a *label plus the real words it stands for*, both in the deck. So
 * two different mistakes are possible and they are not equally bad:
 *
 * - **leaking** — the words reach the model. The prompt path already prevents
 *   this; nothing here can undo a leak that already happened.
 * - **destroying** — the model rewrites the words, or drops the element, and
 *   the reader loses data they chose to keep. Silent, and the file still opens.
 *
 * So this is a damage guard, not a secrecy guard.
 *
 * ## Two shapes of check
 *
 * Ops mostly address an element by id, so a withheld element is simply refused.
 * `findReplace` is the exception: it substitutes text inside runs, and matching
 * there is per-run ("matches that span two differently formatted runs are not
 * replaced"), which makes the test exact and cheap — *if `find` occurs inside a
 * run that is withheld, that match damages the span*. No position arithmetic and
 * no before/after text comparison, which is what the same problem costs in a
 * prose editor.
 *
 * ## What is deliberately still allowed
 *
 * Moving, resizing, recolouring, reordering, restyling and writing speaker notes
 * do not touch a withheld run's text or its mark, and refusing them would make
 * the feature unusable for ordinary work. Turning the underline off is allowed
 * too: the words stay hidden, only the hint goes — the reader's own call.
 */

/**
 * Ops that write into an element's content, so touching a withheld element with
 * one of these costs the reader something. Anything not listed here is geometry,
 * appearance or metadata.
 */
const CONTENT_OPS = new Set([
  'setText',
  'deleteElement',
  'setTableCell',
  'tableMerge',
  'tableStructure',
  'ungroupElement',
  'replacePicture',
  'addMedia',
  'insertEquation',
  'pasteElements',
  'setShapeCustomGeometry',
])

/** Ops that drop a whole page, and so any span on it. */
const PAGE_DELETE_OPS = new Set(['deleteSlide'])

/**
 * The label of whatever this node withholds, if anything — recursing into group
 * children, because a group is a legal op target in its own right and destroying
 * or ungrouping it would take its withheld descendants with it.
 */
function labelOf(node: RenderNode): string | null {
  if (node.type === 'picture') return node.redact ?? null
  if (node.type === 'group') {
    for (const child of node.children) {
      const hit = labelOf(child)
      if (hit) return hit
    }
    return null
  }
  if (node.type !== 'shape' && node.type !== 'text') return null
  for (const line of node.text?.lines ?? []) {
    for (const run of line.runs) if (run.redact) return run.redact
  }
  return null
}

/** The node an op addresses, or null. */
function targetOf(op: Record<string, unknown>, slide: RenderSlide | undefined): RenderNode | null {
  const id = op.elementId ?? (op.target as { el?: unknown } | undefined)?.el
  if (typeof id !== 'string' || !slide) return null
  const walk = (nodes: RenderNode[]): RenderNode | null => {
    for (const n of nodes) {
      if (n.id === id || n.sourceId === id) return n
      if (n.type === 'group') {
        const hit = walk(n.children)
        if (hit) return hit
      }
    }
    return null
  }
  return walk(slide.nodes)
}

/** Any withheld span anywhere on this page, label included. */
function withheldOnSlide(slide: RenderSlide | undefined): string | null {
  if (!slide) return null
  const walk = (nodes: RenderNode[]): string | null => {
    for (const n of nodes) {
      if (n.type === 'group') {
        const hit = walk(n.children)
        if (hit) return hit
        continue
      }
      const label = labelOf(n)
      if (label) return label
    }
    return null
  }
  return walk(slide.nodes)
}

function pageOf(slideIndex: unknown, slides: RenderSlide[]): RenderSlide | undefined {
  if (typeof slideIndex === 'number') return slides[slideIndex]
  return undefined
}

/**
 * Check one op. Returns null when it may proceed, or a failure to hand back.
 *
 * `find` is passed in because only `findReplace` needs it, and threading an
 * optional field through every op would make the refusal reason worse.
 */
export function redactGuardFor(
  op: Record<string, unknown>,
  slides: RenderSlide[],
  find = typeof op.find === 'string' ? op.find : undefined,
): { reason: string; label: string } | null {
  const name = String(op.op ?? '')
  if (!name) return null

  if (name === 'findReplace') {
    if (!find) return null
    // The haystack keeps its case when the search does, so the needle must too.
    // Lowercasing it unconditionally made a case-sensitive search unable to match
    // anything with a capital letter — which silently under-refuses.
    const caseSensitive = op.matchCase === true
    const needle = caseSensitive ? find : find.toLowerCase()
    // matching is per run, so a hit inside a withheld run is exactly the damage
    const walk = (nodes: RenderNode[]): string | null => {
      for (const n of nodes) {
        if (n.type === 'group') {
          const hit = walk(n.children)
          if (hit) return hit
          continue
        }
        if (n.type === 'picture' && n.redact) return n.redact
        if (n.type !== 'shape' && n.type !== 'text') continue
        for (const line of n.text?.lines ?? []) {
          for (const run of line.runs) {
            if (!run.redact) continue
            const hay = caseSensitive ? run.text : run.text.toLowerCase()
            if (hay.includes(needle)) return run.redact
          }
        }
      }
      return null
    }
    for (const slide of slides) {
      const label = walk(slide.nodes)
      if (label) {
        return {
          label,
          reason:
            `this replacement would rewrite text withheld from the model ({{${label}}}). ` +
            `The model was never shown those words, so it cannot have meant to change them. ` +
            `Rephrase around the placeholder instead.`,
        }
      }
    }
    return null
  }

  if (PAGE_DELETE_OPS.has(name)) {
    const label = withheldOnSlide(pageOf(op.slideIndex, slides))
    if (label) {
      return {
        label,
        reason: `deleting this page would drop a span withheld from the model ({{${label}}}). Clear the mark first if you really mean to remove it.`,
      }
    }
    return null
  }

  if (CONTENT_OPS.has(name)) {
    const label = labelOf(targetOf(op, pageOf(op.slideIndex, slides)) ?? ({} as RenderNode))
    if (label) {
      return {
        label,
        reason: `this edit would overwrite a span withheld from the model ({{${label}}}). Its words stay in the deck; rephrase around the placeholder instead of replacing them.`,
      }
    }
  }

  return null
}

/** The marker as the model knows it, for an error message. */
export function guardPlaceholderLabel(label: string): string {
  return placeholderSource(label)
}

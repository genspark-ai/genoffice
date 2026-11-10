/**
 * Paste options for foreign (non-ProseMirror) HTML payloads: web pages, Word,
 * mail clients. Plain Ctrl+V keeps the source formatting (Word's factory
 * default); this module adds the two other Word paste modes and the plumbing
 * the floating post-paste chip and the ribbon Paste button share:
 *
 *  - 'source'  keep source formatting (today's behavior, the default)
 *  - 'merge'   merge formatting: keep emphasis (bold/italic/underline/links)
 *              and block structure, but text formats like typing at the
 *              insertion point (caret docTextStyle + paragraph format)
 *  - 'text'    keep text only: route through the plain-text lane
 *
 * The mode of a paste is decided when the paste EVENT arrives — ProseMirror
 * parses the clipboard before handlePaste runs, so transformPasted cannot ask
 * the event (same lesson as the r181 foreign-paste flag, which this replaces).
 */
import type {
  Fragment as PmFragment,
  Mark,
  Node as PmNode,
  ResolvedPos,
  Schema,
} from '@tiptap/pm/model'
import { Fragment } from '@tiptap/pm/model'
import type { EditorState, SelectionBookmark, Transaction } from '@tiptap/pm/state'
import type { Step } from '@tiptap/pm/transform'
import { PARA_FORMAT_ATTRS } from './caret-marks'
import { isForeignPasteHtml } from './paste-web-html'
import { TRACK_IGNORE } from './revisions'

export type PasteMode = 'source' | 'merge' | 'text'

const MODE_KEY = 'aidocs.pasteFromOtherApps'

/** The user's default mode for pasting from other programs (Word's
 *  "Pasting from other programs" option; factory default = keep source). */
export function defaultPasteMode(): PasteMode {
  const stored = localStorage.getItem(MODE_KEY)
  return stored === 'merge' || stored === 'text' ? stored : 'source'
}

export function setDefaultPasteMode(mode: PasteMode): void {
  localStorage.setItem(MODE_KEY, mode)
}

// ── per-paste handshake ────────────────────────────────────────────────────
// beginForeignPaste (DOM paste handler / ribbon button / chip re-apply) →
// transformPasted consumes the mode → the handlePaste lanes read the consumed
// mode (they run after transformPasted and re-parse or re-route themselves).

let pendingMode: PasteMode | null = null
let overrideMode: PasteMode | null = null
let consumedMode: PasteMode | null = null
let rerouteToText = false

/** Chip re-applies force their mode for exactly the next foreign paste. */
export function forceNextPasteMode(mode: PasteMode): void {
  overrideMode = mode
}

/** Called with the clipboard HTML when a paste event arrives. Returns true
 *  (and arms the per-paste mode) only for foreign fragments — our own
 *  clipboard HTML keeps its exact semantics, drops never arm anything. */
export function beginForeignPaste(html: string): boolean {
  consumedMode = null
  rerouteToText = false
  if (!html || !isForeignPasteHtml(html)) {
    pendingMode = null
    overrideMode = null
    return false
  }
  pendingMode = overrideMode ?? defaultPasteMode()
  overrideMode = null
  return true
}

/** transformPasted: take the armed mode (once per paste). */
export function consumeForeignPaste(): PasteMode | null {
  const mode = pendingMode
  pendingMode = null
  consumedMode = mode
  if (mode === 'text') rerouteToText = true
  return mode
}

/** The mode transformPasted consumed for the paste handlePaste is handling —
 *  the wholesale empty-paragraph lane re-parses the HTML itself and must
 *  apply the same mode. */
export function lastForeignPasteMode(): PasteMode | null {
  return consumedMode
}

/** handlePaste: a text-mode foreign paste diverts to the plain-text lane. */
export function takeTextReroute(): boolean {
  const flag = rerouteToText
  rerouteToText = false
  return flag
}

// ── caret context ──────────────────────────────────────────────────────────

/** The insertion point's docTextStyle mark — the exact style typing there
 *  would produce (storedMarks precedence handled by the caller). */
export function caretStyleMark(marks: readonly Mark[], schema: Schema): Mark | null {
  const type = schema.marks.docTextStyle
  return type ? (marks.find((mark) => mark.type === type) ?? null) : null
}

/** The insertion paragraph's formatting attrs (r178 allowlist — identity and
 *  annotation attrs never clone), or null outside any paragraph. */
export function caretParaFormat($pos: ResolvedPos, schema: Schema): Record<string, unknown> | null {
  const paragraph = schema.nodes.docParagraph
  if (!paragraph) return null
  for (let depth = $pos.depth; depth >= 1; depth--) {
    const node = $pos.node(depth)
    if (node.type === paragraph) {
      const picked: Record<string, unknown> = {}
      for (const key of PARA_FORMAT_ATTRS) if (key in node.attrs) picked[key] = node.attrs[key]
      return picked
    }
  }
  return null
}

// ── merge formatting ───────────────────────────────────────────────────────

/**
 * Word's Merge Formatting: pasted text formats like typing at the insertion
 * point — every run's docTextStyle is replaced by the caret's (or removed,
 * so the text follows the document default exactly like the surrounding
 * text), every paragraph takes the insertion paragraph's format attrs.
 * Emphasis marks (bold/italic/underline/links…) and block structure stay;
 * list/numbering attrs are outside the allowlist and survive. Table subtrees
 * are left alone, matching the source-mode pushdown/fill.
 */
export function mergeFormattingFragment(
  fragment: PmFragment,
  schema: Schema,
  style: Mark | null,
  paraAttrs: Record<string, unknown> | null,
): PmFragment {
  const styleType = schema.marks.docTextStyle
  const paragraph = schema.nodes.docParagraph
  const mapped: PmNode[] = []
  fragment.forEach((node) => {
    if (node.type.spec.tableRole === 'table') {
      mapped.push(node)
      return
    }
    if (node.isText) {
      let marks = styleType ? node.marks.filter((mark) => mark.type !== styleType) : node.marks
      if (style) marks = style.addToSet(marks)
      mapped.push(node.mark(marks))
      return
    }
    const content = mergeFormattingFragment(node.content, schema, style, paraAttrs)
    if (node.type === paragraph && paraAttrs) {
      mapped.push(node.type.create({ ...node.attrs, ...paraAttrs }, content, node.marks))
    } else {
      mapped.push(node.copy(content))
    }
  })
  return Fragment.from(mapped)
}

// ── post-paste chip payload ────────────────────────────────────────────────

export type PastePayload = { html: string; text: string; mode: PasteMode }

let chipPayload: PastePayload | null = null

/** Stashed by the paste entry points; the App transaction listener takes it
 *  when the paste transaction lands and shows the chip. Payloads without
 *  prose (image-only fragments) never stash — there is nothing to re-format. */
export function stashPastePayload(payload: PastePayload | null): void {
  if (payload) {
    try {
      const body = new window.DOMParser().parseFromString(payload.html, 'text/html').body
      if (!(body.textContent ?? '').trim()) payload = null
    } catch {
      payload = null
    }
  }
  chipPayload = payload
}

export function takePastePayload(): PastePayload | null {
  const payload = chipPayload
  chipPayload = null
  return payload
}

// ── chip re-apply ──────────────────────────────────────────────────────────
// A mode change must see the insertion point the ORIGINAL paste saw — caret
// marks, paragraph format, the empty-paragraph wholesale lane — not the
// pasted content. Re-pasting over the pasted range reads the source
// formatting back (merge/text no-op) and never qualifies for the wholesale
// lane (a heading re-applied onto its own text gets demoted). So a re-apply
// undoes the paste, puts the caret back exactly as it was, and pastes again.

export type PasteRestore = {
  /** inverse of the paste's steps, in application order */
  steps: Step[]
  selection: SelectionBookmark
  storedMarks: readonly Mark[] | null
}

/** Snapshot taken when the paste transaction lands: `before` is the state
 *  the paste started from, `transactions` the paste plus its appended
 *  transactions. Null when the two disagree (the paste did not start from
 *  the state we saw). */
export function capturePasteRestore(
  before: EditorState,
  transactions: readonly Transaction[],
): PasteRestore | null {
  if (transactions[0]?.before !== before.doc) return null
  const steps: Step[] = []
  for (const tr of transactions) {
    tr.steps.forEach((step, index) => steps.push(step.invert(tr.docs[index]!)))
  }
  steps.reverse()
  return { steps, selection: before.selection.getBookmark(), storedMarks: before.storedMarks }
}

/** Undo the paste and restore the insertion point (selection and pending
 *  marks), or null when the document no longer accepts the inverse steps. */
export function revertPaste(state: EditorState, restore: PasteRestore): Transaction | null {
  const tr = state.tr
  try {
    for (const step of restore.steps) if (tr.maybeStep(step).failed) return null
    // Taking back our own paste is not an authored deletion; the re-paste that
    // follows in the same tick joins this step in one undo event.
    return tr
      .setSelection(restore.selection.resolve(tr.doc))
      .setStoredMarks(restore.storedMarks)
      .setMeta(TRACK_IGNORE, true)
  } catch {
    return null
  }
}

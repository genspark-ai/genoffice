/**
 * Withheld content for the deck editor: slides the reader keeps but does not
 * want a model to read.
 *
 * A withheld run is an ordinary `TextRun` holding the real words, not a
 * replacement for them. The deck and the .pptx on disk still hold the actual
 * text — losing the reader's own data to hide it from a model would be the
 * worse failure. What changes is the *model's* view: when an element is
 * serialized for a request, a withheld run is written as `{{label}}` instead of
 * its contents. The mark itself is carried by the engine
 * (`TextRun.redact` / `ElementBase.redact`, written to `<a:extLst>` on save and
 * read back on parse), so this module never touches the file.
 *
 * Two consequences follow from the mark sitting on real content:
 *
 * - the model can split it, the way it could not split an atom. A model
 *   breaking one marker across a word and writing
 *   `{{\u5ba2}}\u6237{{\u7535}}\u8bdd` is plausible, so every model edit is
 *   checked before it is allowed into the deck.
 * - a picture can be withheld too, and so can a video or an audio clip. Media
 *   carries no text, so there is nothing to redact — the point is that the
 *   model is not handed a reference it could fetch. That is a stronger claim in
 *   a deck than in a document: `PictureElement.media.target` is a zip path for
 *   an embedded clip but an **external URL** for a linked one, and a signed URL
 *   emitted to a model is a replayable credential. See `mediaModelView`.
 */

import {
  placeholderInstruction as buildInstruction,
  placeholderSource,
  type PlaceholderSpec,
} from '@genoffice/agent-core/redact-core'

/** What this editor's markers need that another's do not. */
const SLIDES_SPEC: PlaceholderSpec = {
  subject: 'presentation',
  opening:
    'Each stands in for something the reader has deliberately withheld from you; you cannot see what is inside, and that is the point.',
  boundary: 'a line break',
  middle: [
    'Write the prose around them as if each stood for the words it replaces, so "call {{客户电话}}" reads as a natural instruction to phone someone.',
    'If a request needs what a placeholder hides, write around it rather than guessing.',
    'A placeholder may stand in for withheld media: a picture, a video or an audio clip. Treat it as content you were not given and write around it — you cannot open it, fetch it, or infer anything from how long or how large it is.',
  ],
}

export function placeholderInstruction(labels: readonly string[]): string {
  return buildInstruction(labels, SLIDES_SPEC)
}

// ---------------------------------------------------------------------------
// Media: withheld pictures, video and audio
// ---------------------------------------------------------------------------

/** What the model is allowed to conclude about a media node. */
export type MediaKind = 'picture' | 'video' | 'audio'

/**
 * The fields of a picture element this module reads. Declared structurally and
 * with an index signature on purpose: an engine `PictureElement` satisfies it,
 * and so does a hand-built literal, which is what lets a test add the fields
 * that must *not* escape (a duration, a description) and watch them dropped.
 */
export interface MediaNodeLike {
  type?: string
  /** `ElementBase.redact` — the engine's element-level withholding mark */
  redact?: string
  name?: string
  media?: { kind: 'video' | 'audio'; target?: string; external?: boolean }
  transform?: { x: number; y: number; w: number; h: number }
  [key: string]: unknown
}

/** The model-facing view of a media node: a whitelist, never the node itself. */
export interface MediaModelView {
  kind: MediaKind
  /** The shape's own name, for the model to refer to it in prose. */
  name?: string
  /** The reader's label for a withheld node. Present only when withheld. */
  label?: string
  /** `{{label}}` for a withheld node — the only handle the model gets. */
  placeholder?: string
  /**
   * The frame, in EMU. Omitted when withheld: a confidential recording's length
   * and size are themselves information about it, and the app already computes
   * video size elsewhere (`src/main/video-size.ts`) so the omission costs
   * nothing downstream.
   */
  frame?: { x: number; y: number; w: number; h: number }
  /**
   * The media's own reference: a zip path when embedded, an **external URL**
   * when linked. Present only when *not* withheld. A withheld node must never
   * carry it — for a linked clip it is a URL the model could fetch, and a signed
   * one is a credential it could replay.
   */
  target?: string
  external?: boolean
}

/** A media node the reader has withheld. */
export function isWithheldMedia(node: MediaNodeLike): boolean {
  return typeof node.redact === 'string' && node.redact !== ''
}

/** The label on a withheld media node, or null when it is not withheld. */
export function mediaLabelOf(node: MediaNodeLike): string | null {
  return isWithheldMedia(node) ? (node.redact as string) : null
}

function mediaKindOf(node: MediaNodeLike): MediaKind {
  return node.media?.kind ?? 'picture'
}

/**
 * What the model may see about a media node.
 *
 * Built as a **whitelist** rather than by deleting fields from a copy, on
 * purpose. A blacklist has to be correct about every field that exists today
 * *and* every field added later, and a field nobody thought of leaks silently
 * — the failure is invisible in the output and only shows up as someone
 * exfiltrating a signed URL. Constructing the view from named fields makes the
 * safe case the default: a new field on `PictureElement` is withheld by
 * construction until somebody deliberately adds it here.
 *
 * That is what drops `mediaRef` and `dataUrl` (a picture's zip path and its
 * decoded bytes, either of which the model could fetch), the external
 * `media.target`, `descr` (free-form `<p:cNvPr>` metadata passed through
 * verbatim by the engine), and any duration a caller may have attached.
 */
export function mediaModelView(node: MediaNodeLike): MediaModelView {
  const kind = mediaKindOf(node)
  const label = mediaLabelOf(node)
  if (label !== null) {
    // The label and nothing else. No target, no frame, no name from the deck's
    // own vocabulary beyond what the reader chose to call it.
    return { kind, label, placeholder: placeholderSource(label) }
  }
  const view: MediaModelView = { kind }
  if (node.name) view.name = node.name
  if (node.transform) view.frame = { ...node.transform }
  if (node.media?.target) view.target = node.media.target
  if (node.media?.external) view.external = true
  return view
}

export {
  MAX_LABEL_LENGTH,
  checkPlaceholders,
  collectPlaceholders,
  isWholePlaceholder,
  placeholderSource,
  readPlaceholderLabel,
  sanitizeLabel,
  type PlaceholderIssue,
} from '@genoffice/agent-core/redact-core'

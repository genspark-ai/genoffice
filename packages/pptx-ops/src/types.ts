import type { NamedAction } from '@genoffice/pptx-engine'

// Edit payload types shared by the op layer and the slides app IPC surface.

/** One rich-text run (sent by the editor, with independent formatting). */
export interface EditRun {
  text: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontSize?: number
  fontFamily?: string
  color?: string
  /** Strikethrough (DOM-authoritative boolean, like bold/italic/underline) */
  strike?: boolean
  /** Super/subscript baseline % (positive = superscript; 0 = none, used to disable explicitly) */
  baseline?: number
  /** Text outline (for WordArt), width in EMU */
  outline?: { color: string; widthEmu: number }
  /** Dynamic field (slidenum / datetime1…); text is the cached value */
  field?: string
  /** Source model run index (index into the original paragraph's runs); the main process uses it to backtrack unedited format fields */
  srcRun?: number
  /** Run hyperlink. undefined = keep the original run's link (programmatic paths that can't
   * express links); null = explicitly none (the editor DOM is authoritative, link removed) */
  link?: LinkTargetOp | null
}

/** One paragraph (with alignment). */
export interface EditParagraph {
  runs: EditRun[]
  align?: 'left' | 'center' | 'right' | 'justify'
  /** Indent level 0..8 (returned after editor Tab/⇧Tab adjustment; defaults to the original paragraph's) */
  level?: number
  /** Source model paragraph index; the main process uses it to inherit bullet/line spacing etc. (both halves of a split share a source) */
  srcPara?: number
  /** Per-paragraph format explicitly changed during this edit session (absent = keep the original) */
  bullet?: 'char' | 'number' | 'blip' | 'none'
  bulletChar?: string
  bulletFont?: string
  /** buAutoNum scheme / first number (with bullet: 'number', or alone on a numbered paragraph) */
  numType?: string
  startAt?: number
  /** Picture bullet source (with bullet: 'blip'); the op layer lands it as a media part */
  bulletImage?: { base64: string; ext: string }
  lineSpacingPct?: number
  spaceBeforePt?: number
  spaceAfterPt?: number
  /** Paragraph base direction toggled during this edit session (false = explicit LTR) */
  rtl?: boolean
}

/** One geometry primitive collected by the edit-script sandbox (px, viewport space). */
export interface ScriptBoxOp {
  id: string
  x: number
  y: number
  w: number
  h: number
  rotation: number
  /** Group child: converted to child-space EMU by the main-process shim */
  groupId?: string
}

/** setStyle's style-override fields (pass only what changes; align is paragraph-level, the rest override per run). */
export interface ScriptStylePatch {
  fontSize?: number
  color?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontFamily?: string
  align?: 'left' | 'center' | 'right'
}

/** One non-geometry primitive collected by the edit-script sandbox, in script call order. */
export type ScriptEditOp = (
  | { kind: 'text'; paragraphs: EditParagraph[] }
  | { kind: 'style'; style: ScriptStylePatch }
  | { kind: 'fill'; fill: string }
  | { kind: 'stroke'; stroke: { color: string; widthPt: number } | null }
) & { id: string; groupId?: string }

export interface ApplyEditScriptOp {
  slideIndex: number
  fitWidthPx: number
  boxes: ScriptBoxOp[]
  edits: ScriptEditOp[]
}

/** Element hyperlink target. */
export type LinkTargetOp =
  | { kind: 'url'; url: string }
  | { kind: 'slide'; slideIndex: number }
  | { kind: 'action'; action: NamedAction }

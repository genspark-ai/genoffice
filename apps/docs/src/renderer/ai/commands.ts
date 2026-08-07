import type { Editor } from '@tiptap/core'
import type { Node as PmDocNode, Schema } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import { generateTocFieldXml, type TocEntry } from '@genoffice/docx-engine'
import { isEastAsianFontName } from '../font-list'
import { t } from '../i18n/locale'
import { isTrackedDeleted, liveText } from './protocol'

/**
 * Structured edit commands: the model emits a
 * batchUpdate-style JSON envelope and code applies it deterministically on the
 * PM doc — one transaction per envelope, so one undo reverts the whole round.
 * The docx round trip is untouched: changed nodes become dirty via the normal
 * signature comparison and regenerate on save; docxIndex anchors are never
 * modified here.
 */

// ---- command types (§4) ----

/** All given conditions filter with AND; at least one condition is required */
export interface Target {
  /** 'image' matches protected image blocks (docProtected + blockType image) */
  nodeType?: 'docHeading' | 'docParagraph' | 'docListItem' | 'image'
  /** Restrict heading level (only with nodeType: 'docHeading') */
  headingLevel?: number
  /** The block's plain text contains this substring (case-sensitive unless matchCase: false) */
  containsText?: string
  matchCase?: boolean
  /** By block index (the index output by buildDocumentContext, 0-based, current PM doc top-level order) */
  blockIndexes?: number[]
  /** 'selection' = only blocks covered by the current selection; default = whole document */
  scope?: 'selection' | 'document'
}

/** Run-level style; only keys listed in fields take effect, and null clears the property */
export interface UpdateTextStyle {
  target: Target
  style: {
    color?: string | null
    highlight?: string | null
    sizeHalfPoints?: number | null
    font?: string | null
    bold?: boolean
    italic?: boolean
    underline?: boolean
    strike?: boolean
    /** Super/subscript (enum values aligned with Google TextStyle.baselineOffset) */
    baselineOffset?: 'SUPERSCRIPT' | 'SUBSCRIPT' | 'NONE' | null
    /** Hyperlink (aligned with Google TextStyle.link); null removes the link */
    link?: { url: string } | null
  }
  fields: Array<keyof UpdateTextStyle['style']>
}

/** Paragraph-level format; likewise constrained by the fields mask */
export interface UpdateParagraphStyle {
  target: Target
  style: {
    align?: 'left' | 'center' | 'right' | 'justify' | null
    lineSpacing?: number | null
    indentLeft?: number | null
    indentRight?: number | null
    /** First-line indent in twips; positive = first-line indent, negative = hanging indent (aligned with Google indentFirstLine) */
    indentFirstLine?: number | null
    spaceBefore?: number | null
    spaceAfter?: number | null
    pageBreakBefore?: boolean
    shadingFill?: string | null
    /** Paragraph borders, a subset of "tblr" (simplified Google borderTop/Bottom/Left/Right) */
    borders?: string | null
  }
  fields: Array<keyof UpdateParagraphStyle['style']>
}

/** Paragraph promote/demote; level 0 = demote to body paragraph, 1-6 = the corresponding heading level */
export interface SetHeadingLevel {
  target: Target
  level: 0 | 1 | 2 | 3 | 4 | 5 | 6
}

/** Whole-document find & replace (plain text, no cross-block matching) */
export interface ReplaceAllText {
  containsText: string
  replaceText: string
  matchCase?: boolean
}

export interface DeleteBlocks {
  target: Target
}

/** Move the blocks in blockIndexes (relative order kept) after afterBlockIndex; -1 = document start */
export interface MoveBlocks {
  blockIndexes: number[]
  afterBlockIndex: number
}

/**
 * Paragraphs to list (name and semantics aligned with Google's createParagraphBullets).
 * BULLET*-prefixed presets = unordered, NUMBERED* = ordered; heading blocks aren't converted (counted in matched but unchanged).
 */
export interface CreateParagraphBullets {
  target: Target
  bulletPreset?: string
}

/** List items back to body paragraphs (aligned with Google's deleteParagraphBullets) */
export interface DeleteParagraphBullets {
  target: Target
}

/** Image block properties; giving only one of widthPx/heightPx scales the other side by the original ratio */
export interface UpdateImageProperties {
  target: Target
  properties: {
    widthPx?: number | null
    heightPx?: number | null
    align?: 'left' | 'center' | 'right' | null
  }
  fields: Array<keyof UpdateImageProperties['properties']>
}

/**
 * Insert a TOC field after the given block (a real TOC field; Word recomputes page numbers on open).
 * Entries come from the current document's heading blocks; the command fails when there are no headings.
 */
export interface InsertToc {
  /** Insert after this block index; -1 = document start */
  afterBlockIndex: number
}

export type Command =
  | { updateTextStyle: UpdateTextStyle }
  | { updateParagraphStyle: UpdateParagraphStyle }
  | { setHeadingLevel: SetHeadingLevel }
  | { replaceAllText: ReplaceAllText }
  | { deleteBlocks: DeleteBlocks }
  | { moveBlocks: MoveBlocks }
  | { createParagraphBullets: CreateParagraphBullets }
  | { deleteParagraphBullets: DeleteParagraphBullets }
  | { updateImageProperties: UpdateImageProperties }
  | { insertToc: InsertToc }

/** numbering context so converted list items join real docx numbering on save */
export interface CommandContext {
  numIds?: { bullet: string | null; ordered: string | null } | null
  /** record the envelope's changes as tracked revisions under this author */
  track?: { author: string } | null
}

export interface CommandEnvelope {
  commands: Command[]
}

export interface CommandResult {
  command: string
  matched: number
  changed: number
  skippedProtected: number
  /** targets skipped because the text is a pending tracked deletion */
  skippedDeleted?: number
  detail?: string
}

export interface ExecuteOutcome {
  ok: boolean
  results: CommandResult[]
  summary: string
  error?: string
}

// ---- envelope extraction ----

/** pull a command envelope out of a model response; null = not command mode */
export function parseCommandEnvelope(raw: string): CommandEnvelope | null {
  const text = raw.trim()
  const fence = /```json\s*([\s\S]*?)```/.exec(text)
  const candidate = fence ? fence[1].trim() : text.startsWith('{') ? text : null
  if (!candidate) return null
  try {
    const parsed: unknown = JSON.parse(candidate)
    if (
      parsed &&
      typeof parsed === 'object' &&
      Array.isArray((parsed as CommandEnvelope).commands)
    ) {
      return parsed as CommandEnvelope
    }
  } catch {
    // not JSON: content mode
  }
  return null
}

// ---- validation (whole-envelope: any invalid command rejects everything) ----

const NODE_TYPES = ['docHeading', 'docParagraph', 'docListItem', 'image'] as const
const TEXT_STYLE_KEYS = [
  'color',
  'highlight',
  'sizeHalfPoints',
  'font',
  'bold',
  'italic',
  'underline',
  'strike',
  'baselineOffset',
  'link',
] as const
const PARA_STYLE_KEYS = [
  'align',
  'lineSpacing',
  'indentLeft',
  'indentRight',
  'indentFirstLine',
  'spaceBefore',
  'spaceAfter',
  'pageBreakBefore',
  'shadingFill',
  'borders',
] as const
const BASELINE_OFFSETS = ['SUPERSCRIPT', 'SUBSCRIPT', 'NONE'] as const
const COMMAND_NAMES = [
  'updateTextStyle',
  'updateParagraphStyle',
  'setHeadingLevel',
  'replaceAllText',
  'deleteBlocks',
  'moveBlocks',
  'createParagraphBullets',
  'deleteParagraphBullets',
  'updateImageProperties',
  'insertToc',
] as const
const IMAGE_PROP_KEYS = ['widthPx', 'heightPx', 'align'] as const

type CommandName = (typeof COMMAND_NAMES)[number]

function validateTarget(target: unknown, where: string): string | null {
  if (!target || typeof target !== 'object') return `${where}: missing target`
  const t = target as Target
  if (t.nodeType !== undefined && !NODE_TYPES.includes(t.nodeType)) {
    return `${where}: unknown nodeType "${String(t.nodeType)}"`
  }
  if (
    t.headingLevel !== undefined &&
    (!Number.isInteger(t.headingLevel) || t.headingLevel < 1 || t.headingLevel > 6)
  ) {
    return `${where}: headingLevel must be an integer between 1 and 6`
  }
  if (
    t.blockIndexes !== undefined &&
    (!Array.isArray(t.blockIndexes) || t.blockIndexes.some((i) => !Number.isInteger(i) || i < 0))
  ) {
    return `${where}: blockIndexes must be an array of non-negative integers`
  }
  const hasCondition =
    t.nodeType !== undefined ||
    t.headingLevel !== undefined ||
    typeof t.containsText === 'string' ||
    (Array.isArray(t.blockIndexes) && t.blockIndexes.length > 0) ||
    t.scope === 'selection'
  if (!hasCondition) return `${where}: target requires at least one condition`
  return null
}

function validateStyleCommand(
  payload: { target?: unknown; style?: unknown; fields?: unknown },
  allowedKeys: readonly string[],
  where: string,
): string | null {
  const targetError = validateTarget(payload.target, where)
  if (targetError) return targetError
  if (!payload.style || typeof payload.style !== 'object') return `${where}: missing style`
  if (!Array.isArray(payload.fields) || payload.fields.length === 0) {
    return `${where}: fields must not be empty`
  }
  const bad = payload.fields.filter((f) => !allowedKeys.includes(String(f)))
  if (bad.length > 0) return `${where}: unknown fields ${bad.map(String).join(', ')}`
  return null
}

export function validateEnvelope(envelope: unknown): string | null {
  if (!envelope || typeof envelope !== 'object') return 'the envelope must be an object'
  const commands = (envelope as CommandEnvelope).commands
  if (!Array.isArray(commands) || commands.length === 0) return 'commands must be a non-empty array'
  for (let i = 0; i < commands.length; i++) {
    const command = commands[i] as Record<string, unknown>
    if (!command || typeof command !== 'object') return `command #${i}: must be an object`
    const keys = Object.keys(command)
    if (keys.length !== 1) return `command #${i}: must be a single-key object`
    const name = keys[0]
    const where = `command #${i} ${name}`
    if (!COMMAND_NAMES.includes(name as CommandName)) return `${where}: unknown command`
    const payload = command[name] as Record<string, unknown>
    if (!payload || typeof payload !== 'object') return `${where}: parameters must be an object`
    let error: string | null = null
    switch (name as CommandName) {
      case 'updateTextStyle': {
        error = validateStyleCommand(payload, TEXT_STYLE_KEYS, where)
        if (!error) {
          const style = payload.style as UpdateTextStyle['style']
          if (
            style.baselineOffset != null &&
            !(BASELINE_OFFSETS as readonly string[]).includes(style.baselineOffset)
          ) {
            error = `${where}: baselineOffset must be SUPERSCRIPT/SUBSCRIPT/NONE`
          } else if (style.link != null && typeof style.link.url !== 'string') {
            error = `${where}: link must be { url } or null`
          }
        }
        break
      }
      case 'updateParagraphStyle':
        error = validateStyleCommand(payload, PARA_STYLE_KEYS, where)
        break
      case 'setHeadingLevel': {
        error = validateTarget(payload.target, where)
        const level = payload.level
        if (!error && (!Number.isInteger(level) || Number(level) < 0 || Number(level) > 6)) {
          error = `${where}: level must be an integer between 0 and 6`
        }
        break
      }
      case 'replaceAllText':
        if (typeof payload.containsText !== 'string' || payload.containsText === '') {
          error = `${where}: containsText must not be empty`
        } else if (typeof payload.replaceText !== 'string') {
          error = `${where}: replaceText must be a string`
        }
        break
      case 'deleteBlocks':
        error = validateTarget(payload.target, where)
        break
      case 'moveBlocks':
        if (
          !Array.isArray(payload.blockIndexes) ||
          payload.blockIndexes.length === 0 ||
          payload.blockIndexes.some((i) => !Number.isInteger(i) || Number(i) < 0)
        ) {
          error = `${where}: blockIndexes must be a non-empty array of non-negative integers`
        } else if (
          !Number.isInteger(payload.afterBlockIndex) ||
          Number(payload.afterBlockIndex) < -1
        ) {
          error = `${where}: afterBlockIndex must be an integer ≥ -1`
        }
        break
      case 'createParagraphBullets': {
        error = validateTarget(payload.target, where)
        const preset = payload.bulletPreset
        if (
          !error &&
          preset !== undefined &&
          (typeof preset !== 'string' || !/^(BULLET|NUMBERED)/.test(preset))
        ) {
          error = `${where}: bulletPreset must start with BULLET or NUMBERED`
        }
        break
      }
      case 'deleteParagraphBullets':
        error = validateTarget(payload.target, where)
        break
      case 'updateImageProperties': {
        error = validateTarget(payload.target, where)
        if (!error) {
          if (!payload.properties || typeof payload.properties !== 'object') {
            error = `${where}: missing properties`
          } else if (!Array.isArray(payload.fields) || payload.fields.length === 0) {
            error = `${where}: fields must not be empty`
          } else {
            const bad = payload.fields.filter(
              (f) => !(IMAGE_PROP_KEYS as readonly string[]).includes(String(f)),
            )
            if (bad.length > 0) error = `${where}: unknown fields ${bad.map(String).join(', ')}`
          }
        }
        break
      }
      case 'insertToc':
        if (!Number.isInteger(payload.afterBlockIndex) || Number(payload.afterBlockIndex) < -1) {
          error = `${where}: afterBlockIndex must be an integer ≥ -1`
        }
        break
    }
    if (error) return error
  }
  return null
}

// ---- target matching ----

interface TopBlock {
  index: number
  pos: number
  node: PmDocNode
}

interface SelRange {
  from: number
  to: number
}

function topLevelBlocks(doc: PmDocNode): TopBlock[] {
  const out: TopBlock[] = []
  let index = 0
  doc.forEach((node, offset) => {
    out.push({ index, pos: offset, node })
    index++
  })
  return out
}

function blockText(node: PmDocNode): string {
  if (node.type.name === 'docProtected') {
    return [node.attrs.label, node.attrs.previewText].filter(Boolean).join(' ')
  }
  // match against current content only, or targets would hit already-deleted revision text
  return liveText(node)
}

function matchTarget(doc: PmDocNode, target: Target, sel: SelRange): TopBlock[] {
  return topLevelBlocks(doc).filter((b) => {
    if (target.blockIndexes && !target.blockIndexes.includes(b.index)) return false
    if (target.nodeType === 'image') {
      if (b.node.type.name !== 'docProtected' || b.node.attrs.blockType !== 'image') return false
    } else if (target.nodeType && b.node.type.name !== target.nodeType) {
      return false
    }
    if (target.headingLevel !== undefined) {
      if (b.node.type.name !== 'docHeading') return false
      if (Number(b.node.attrs.level) !== target.headingLevel) return false
    }
    if (target.containsText) {
      const text = blockText(b.node)
      const hit =
        target.matchCase === false
          ? text.toLowerCase().includes(target.containsText.toLowerCase())
          : text.includes(target.containsText)
      if (!hit) return false
    }
    if (target.scope === 'selection') {
      const from = b.pos
      const to = b.pos + b.node.nodeSize
      const overlaps = to > sel.from && from < sel.to
      const caretInside = sel.from === sel.to && sel.from >= from && sel.from <= to
      if (!overlaps && !caretInside) return false
    }
    return true
  })
}

// ---- executors (each reads tr.doc fresh, so sequential commands compose) ----

const BOOL_MARK_TYPES: Record<string, string> = {
  bold: 'bold',
  italic: 'italic',
  underline: 'underline',
  strike: 'strike',
}
const TEXT_ATTR_KEYS = ['color', 'highlight', 'sizeHalfPoints', 'font'] as const

function markChanged(tr: Transaction, pos: number): boolean {
  const node = tr.doc.nodeAt(pos)
  if (!node) return false
  tr.setNodeMarkup(pos, undefined, { ...node.attrs, aiChanged: true })
  return true
}

function runUpdateTextStyle(
  tr: Transaction,
  schema: Schema,
  cmd: UpdateTextStyle,
  sel: SelRange,
): CommandResult {
  const matched = matchTarget(tr.doc, cmd.target, sel)
  const boolFields = cmd.fields.filter((f) => f in BOOL_MARK_TYPES)
  // patch applied onto each text node's existing docTextStyle attrs
  const markPatch: Record<string, unknown> = {}
  for (const f of cmd.fields) {
    if (f === 'font') {
      // route to the matching rFonts slot; clearing clears both
      const v = cmd.style.font
      if (!v) Object.assign(markPatch, { font: null, fontAscii: null })
      else if (isEastAsianFontName(v)) markPatch.font = v
      else markPatch.fontAscii = v
    } else if ((TEXT_ATTR_KEYS as readonly string[]).includes(f)) {
      markPatch[f] = cmd.style[f as (typeof TEXT_ATTR_KEYS)[number]] ?? null
    } else if (f === 'baselineOffset') {
      const value = cmd.style.baselineOffset
      markPatch.vertAlign =
        value === 'SUPERSCRIPT' ? 'superscript' : value === 'SUBSCRIPT' ? 'subscript' : null
    }
  }
  const hasLink = cmd.fields.includes('link')
  let changed = 0
  let skippedProtected = 0

  for (const b of matched) {
    if (b.node.type.name === 'docProtected') {
      skippedProtected++
      continue
    }
    const before = b.node
    const from = b.pos + 1
    const to = from + b.node.content.size
    if (from < to) {
      for (const f of boolFields) {
        const markType = schema.marks[BOOL_MARK_TYPES[f]]
        if (cmd.style[f]) tr.addMark(from, to, markType.create())
        else tr.removeMark(from, to, markType)
      }
      if (hasLink) {
        if (cmd.style.link?.url) {
          tr.addMark(from, to, schema.marks.link.create({ href: cmd.style.link.url, rId: null }))
        } else {
          tr.removeMark(from, to, schema.marks.link)
        }
      }
      if (Object.keys(markPatch).length > 0) {
        // merge per text node: only fields-listed attrs change, the rest survive
        b.node.forEach((child, offset) => {
          if (!child.isText) return
          const childFrom = from + offset
          const childTo = childFrom + child.nodeSize
          const existing = child.marks.find((m) => m.type.name === 'docTextStyle')?.attrs ?? {}
          // spread keeps every non-listed attr (fontAscii/styleId/rawRPr…) alive
          const merged: Record<string, unknown> = { ...existing, ...markPatch }
          const empty = Object.values(merged).every((v) => v === null)
          if (empty) tr.removeMark(childFrom, childTo, schema.marks.docTextStyle)
          else tr.addMark(childFrom, childTo, schema.marks.docTextStyle.create(merged))
        })
      }
    }
    const after = tr.doc.nodeAt(b.pos)
    if (after && !after.eq(before)) {
      changed++
      markChanged(tr, b.pos)
    }
  }
  return { command: 'updateTextStyle', matched: matched.length, changed, skippedProtected }
}

function runUpdateParagraphStyle(
  tr: Transaction,
  cmd: UpdateParagraphStyle,
  sel: SelRange,
): CommandResult {
  const matched = matchTarget(tr.doc, cmd.target, sel)
  let changed = 0
  let skippedProtected = 0
  for (const b of matched) {
    if (b.node.type.name === 'docProtected') {
      skippedProtected++
      continue
    }
    const attrs: Record<string, unknown> = { ...b.node.attrs }
    for (const f of cmd.fields) {
      attrs[f] = cmd.style[f] ?? (f === 'pageBreakBefore' ? false : null)
    }
    const dirty = cmd.fields.some((f) => attrs[f] !== b.node.attrs[f])
    if (!dirty) continue
    tr.setNodeMarkup(b.pos, undefined, { ...attrs, aiChanged: true })
    changed++
  }
  return { command: 'updateParagraphStyle', matched: matched.length, changed, skippedProtected }
}

function runSetHeadingLevel(
  tr: Transaction,
  schema: Schema,
  cmd: SetHeadingLevel,
  sel: SelRange,
): CommandResult {
  const matched = matchTarget(tr.doc, cmd.target, sel)
  let changed = 0
  let skippedProtected = 0
  for (const b of matched) {
    if (b.node.type.name === 'docProtected') {
      skippedProtected++
      continue
    }
    const isHeading = b.node.type.name === 'docHeading'
    if (cmd.level === 0 && b.node.type.name === 'docParagraph') continue
    if (cmd.level > 0 && isHeading && Number(b.node.attrs.level) === cmd.level) continue
    const type = cmd.level === 0 ? schema.nodes.docParagraph : schema.nodes.docHeading
    // styleId is tied to the old block role; generate.ts picks heading styles by level
    const attrs: Record<string, unknown> = {
      ...b.node.attrs,
      styleId: null,
      aiChanged: true,
    }
    if (cmd.level > 0) attrs.level = cmd.level
    tr.setNodeMarkup(b.pos, type, attrs)
    changed++
  }
  return { command: 'setHeadingLevel', matched: matched.length, changed, skippedProtected }
}

function runReplaceAllText(tr: Transaction, schema: Schema, cmd: ReplaceAllText): CommandResult {
  const matchCase = cmd.matchCase !== false
  const needle = matchCase ? cmd.containsText : cmd.containsText.toLowerCase()
  const blocks = topLevelBlocks(tr.doc)
  const replacements: Array<{ from: number; to: number; marks: PmDocNode['marks'] }> = []
  const touchedIndexes = new Set<number>()
  let skippedProtected = 0
  let skippedDeleted = 0

  for (const b of blocks) {
    if (b.node.type.name === 'docProtected') {
      if (blockText(b.node).includes(cmd.containsText)) skippedProtected++
      continue
    }
    const blockDeleted = isTrackedDeleted(b.node)
    b.node.forEach((child, offset) => {
      if (!child.isText || !child.text) return
      const struck = blockDeleted || child.marks.some((m) => m.type.name === 'del')
      const hay = matchCase ? child.text : child.text.toLowerCase()
      let at = hay.indexOf(needle)
      while (at !== -1) {
        if (struck) {
          skippedDeleted++
        } else {
          const from = b.pos + 1 + offset + at
          replacements.push({ from, to: from + needle.length, marks: child.marks })
          touchedIndexes.add(b.index)
        }
        at = hay.indexOf(needle, at + needle.length)
      }
    })
  }

  // apply back-to-front so earlier positions stay valid
  for (const r of [...replacements].reverse()) {
    if (cmd.replaceText) tr.replaceWith(r.from, r.to, schema.text(cmd.replaceText, r.marks))
    else tr.delete(r.from, r.to)
  }
  if (touchedIndexes.size > 0) {
    for (const b of topLevelBlocks(tr.doc)) {
      if (touchedIndexes.has(b.index)) markChanged(tr, b.pos)
    }
  }
  return {
    command: 'replaceAllText',
    matched: touchedIndexes.size,
    changed: touchedIndexes.size,
    skippedProtected,
    skippedDeleted,
    detail: t('aiCmdReplacedCount', { count: replacements.length }),
  }
}

function runDeleteBlocks(
  tr: Transaction,
  schema: Schema,
  cmd: DeleteBlocks,
  sel: SelRange,
): CommandResult {
  const blocks = topLevelBlocks(tr.doc)
  const matchedAll = matchTarget(tr.doc, cmd.target, sel)
  // deleting a pending deletion is a no-op (the tracker re-materializes it struck);
  // skip and report so the model does not retry forever
  const matched = matchedAll.filter((b) => !isTrackedDeleted(b.node))
  const skippedDeleted = matchedAll.length - matched.length
  if (matched.length === blocks.length && matched.length > 0) {
    // doc requires block+: deleting everything leaves one empty paragraph
    tr.replaceWith(0, tr.doc.content.size, schema.nodes.docParagraph.create())
    return {
      command: 'deleteBlocks',
      matched: matchedAll.length,
      changed: matched.length,
      skippedProtected: 0,
      skippedDeleted,
    }
  }
  for (const b of [...matched].sort((a, z) => z.pos - a.pos)) {
    tr.delete(b.pos, b.pos + b.node.nodeSize)
  }
  return {
    command: 'deleteBlocks',
    matched: matchedAll.length,
    changed: matched.length,
    skippedProtected: 0,
    skippedDeleted,
  }
}

function runMoveBlocks(tr: Transaction, cmd: MoveBlocks): CommandResult {
  const blocks = topLevelBlocks(tr.doc)
  const movedSet = new Set(cmd.blockIndexes)
  const outOfRange = cmd.blockIndexes.filter((i) => i >= blocks.length)
  if (outOfRange.length > 0)
    throw new Error(`moveBlocks: block indexes out of range ${outOfRange.join(', ')}`)
  if (cmd.afterBlockIndex >= blocks.length) {
    throw new Error(`moveBlocks: afterBlockIndex out of range ${cmd.afterBlockIndex}`)
  }
  if (movedSet.has(cmd.afterBlockIndex)) {
    throw new Error('moveBlocks: afterBlockIndex must not be one of the moved blocks')
  }

  // moved blocks keep document order; aiChanged marks the landing highlight
  const moved = [...movedSet]
    .sort((a, z) => a - z)
    .map((i) => {
      const node = blocks[i].node
      if (node.type.name === 'docProtected') return node
      return node.type.create({ ...node.attrs, aiChanged: true }, node.content, node.marks)
    })
  const rest = blocks.filter((b) => !movedSet.has(b.index))
  let insertAt = 0
  if (cmd.afterBlockIndex >= 0) {
    insertAt = rest.findIndex((b) => b.index === cmd.afterBlockIndex) + 1
  }
  const children = [
    ...rest.slice(0, insertAt).map((b) => b.node),
    ...moved,
    ...rest.slice(insertAt).map((b) => b.node),
  ]
  tr.replaceWith(0, tr.doc.content.size, children)
  return {
    command: 'moveBlocks',
    matched: moved.length,
    changed: moved.length,
    skippedProtected: 0,
  }
}

function runCreateParagraphBullets(
  tr: Transaction,
  schema: Schema,
  cmd: CreateParagraphBullets,
  sel: SelRange,
  context: CommandContext,
): CommandResult {
  const matched = matchTarget(tr.doc, cmd.target, sel)
  const kind = cmd.bulletPreset?.startsWith('NUMBERED') ? 'ordered' : 'bullet'
  const numId = context.numIds?.[kind] ?? null
  let changed = 0
  let skippedProtected = 0
  for (const b of matched) {
    if (b.node.type.name === 'docProtected') {
      skippedProtected++
      continue
    }
    // headings keep their role; converting them to list items would drop the level
    if (b.node.type.name === 'docHeading') continue
    if (b.node.type.name === 'docListItem') {
      if (b.node.attrs.kind === kind) continue
      tr.setNodeMarkup(b.pos, undefined, {
        ...b.node.attrs,
        kind,
        numId: numId ?? b.node.attrs.numId,
        aiChanged: true,
      })
      changed++
      continue
    }
    tr.setNodeMarkup(b.pos, schema.nodes.docListItem, {
      ...b.node.attrs,
      styleId: null,
      kind,
      numId,
      ilvl: 0,
      aiChanged: true,
    })
    changed++
  }
  return { command: 'createParagraphBullets', matched: matched.length, changed, skippedProtected }
}

function runDeleteParagraphBullets(
  tr: Transaction,
  schema: Schema,
  cmd: DeleteParagraphBullets,
  sel: SelRange,
): CommandResult {
  const matched = matchTarget(tr.doc, cmd.target, sel)
  let changed = 0
  let skippedProtected = 0
  for (const b of matched) {
    if (b.node.type.name === 'docProtected') {
      skippedProtected++
      continue
    }
    if (b.node.type.name !== 'docListItem') continue
    tr.setNodeMarkup(b.pos, schema.nodes.docParagraph, {
      ...b.node.attrs,
      styleId: null,
      aiChanged: true,
    })
    changed++
  }
  return { command: 'deleteParagraphBullets', matched: matched.length, changed, skippedProtected }
}

function runUpdateImageProperties(
  tr: Transaction,
  cmd: UpdateImageProperties,
  sel: SelRange,
): CommandResult {
  const matched = matchTarget(tr.doc, cmd.target, sel)
  let changed = 0
  for (const b of matched) {
    if (b.node.type.name !== 'docProtected' || b.node.attrs.blockType !== 'image') continue
    const attrs: Record<string, unknown> = { ...b.node.attrs }
    if (cmd.fields.includes('align')) attrs.imageAlign = cmd.properties.align ?? null
    const wantW = cmd.fields.includes('widthPx') ? (cmd.properties.widthPx ?? null) : null
    const wantH = cmd.fields.includes('heightPx') ? (cmd.properties.heightPx ?? null) : null
    const curW = Number(b.node.attrs.imageWidthPx) || null
    const curH = Number(b.node.attrs.imageHeightPx) || null
    if (wantW && wantH) {
      attrs.imageWidthPx = Math.round(wantW)
      attrs.imageHeightPx = Math.round(wantH)
    } else if (wantW && curW && curH) {
      attrs.imageWidthPx = Math.round(wantW)
      attrs.imageHeightPx = Math.round((wantW * curH) / curW)
    } else if (wantH && curW && curH) {
      attrs.imageHeightPx = Math.round(wantH)
      attrs.imageWidthPx = Math.round((wantH * curW) / curH)
    }
    // save-side patching needs both dimensions; a lone dimension with unknown
    // original size cannot be scaled and is skipped
    const dirty = Object.keys(attrs).some((k) => attrs[k] !== b.node.attrs[k])
    if (!dirty) continue
    tr.setNodeMarkup(b.pos, undefined, attrs)
    changed++
  }
  return { command: 'updateImageProperties', matched: matched.length, changed, skippedProtected: 0 }
}

function runInsertToc(tr: Transaction, schema: Schema, cmd: InsertToc): CommandResult {
  const blocks = topLevelBlocks(tr.doc)
  const entries: TocEntry[] = []
  for (const b of blocks) {
    if (b.node.type.name === 'docHeading' && b.node.textContent.trim()) {
      entries.push({ level: Number(b.node.attrs.level) || 1, text: b.node.textContent })
    }
  }
  if (entries.length === 0) {
    throw new Error(
      'insertToc: the document has no headings (levels 1-6), cannot generate a table of contents',
    )
  }
  if (cmd.afterBlockIndex >= blocks.length) {
    throw new Error(`insertToc: afterBlockIndex out of range ${cmd.afterBlockIndex}`)
  }
  // one protected block per TOC line (same shape as the ribbon's TOC button);
  // the begin fldChar is dirty so Word recalculates page numbers on open
  const nodes = generateTocFieldXml(entries).map((xml, i) =>
    schema.nodes.docProtected.create({
      docxIndex: null,
      blockType: 'passthrough',
      label: 'TOC field',
      genXml: xml,
      fieldDisplay: { kind: 'tocLine', left: entries[i].text, right: '', level: entries[i].level },
    }),
  )
  const pos =
    cmd.afterBlockIndex === -1
      ? 0
      : blocks[cmd.afterBlockIndex].pos + blocks[cmd.afterBlockIndex].node.nodeSize
  tr.insert(pos, nodes)
  return {
    command: 'insertToc',
    matched: entries.length,
    changed: nodes.length,
    skippedProtected: 0,
  }
}

// ---- entry point ----

function summarize(results: CommandResult[]): string {
  const total = results.reduce((sum, r) => sum + r.changed, 0)
  if (total === 0) {
    const skipped = results.reduce((sum, r) => sum + r.skippedProtected, 0)
    return skipped > 0 ? t('aiCmdNoneSkipped', { count: skipped }) : t('aiCmdNone')
  }
  const parts = results.map((r) => {
    let part: string
    switch (r.command) {
      case 'updateTextStyle':
        part = t('aiCmdTextStyle', { count: r.changed })
        break
      case 'updateParagraphStyle':
        part = t('aiCmdParaStyle', { count: r.changed })
        break
      case 'setHeadingLevel':
        part = t('aiCmdHeading', { count: r.changed })
        break
      case 'replaceAllText':
        part = t('aiCmdReplaceAll', {
          count: r.changed,
          detail: r.detail ?? t('aiCmdReplaceFallback'),
        })
        break
      case 'deleteBlocks':
        part = t('aiCmdDeleted', { count: r.changed })
        break
      case 'moveBlocks':
        part = t('aiCmdMoved', { count: r.changed })
        break
      case 'createParagraphBullets':
        part = t('aiCmdToList', { count: r.changed })
        break
      case 'deleteParagraphBullets':
        part = t('aiCmdToBody', { count: r.changed })
        break
      case 'updateImageProperties':
        part = t('aiCmdImages', { count: r.changed })
        break
      case 'insertToc':
        part = t('aiCmdToc', { count: r.matched })
        break
      default:
        part = `${r.command}: ${r.changed}`
    }
    if (r.skippedProtected > 0) part += t('aiCmdSkipped', { count: r.skippedProtected })
    return part
  })
  return parts.join(';')
}

export function executeCommands(
  editor: Editor,
  envelope: CommandEnvelope,
  context: CommandContext = {},
): ExecuteOutcome {
  const validationError = validateEnvelope(envelope)
  if (validationError) return { ok: false, results: [], summary: '', error: validationError }

  const tr = editor.state.tr
  const schema = editor.schema
  const selection = editor.state.selection
  const results: CommandResult[] = []
  try {
    for (const command of envelope.commands) {
      // selection positions mapped through the steps applied so far
      const sel: SelRange = {
        from: tr.mapping.map(selection.from),
        to: tr.mapping.map(selection.to),
      }
      if ('updateTextStyle' in command) {
        results.push(runUpdateTextStyle(tr, schema, command.updateTextStyle, sel))
      } else if ('updateParagraphStyle' in command) {
        results.push(runUpdateParagraphStyle(tr, command.updateParagraphStyle, sel))
      } else if ('setHeadingLevel' in command) {
        results.push(runSetHeadingLevel(tr, schema, command.setHeadingLevel, sel))
      } else if ('replaceAllText' in command) {
        results.push(runReplaceAllText(tr, schema, command.replaceAllText))
      } else if ('deleteBlocks' in command) {
        results.push(runDeleteBlocks(tr, schema, command.deleteBlocks, sel))
      } else if ('createParagraphBullets' in command) {
        results.push(
          runCreateParagraphBullets(tr, schema, command.createParagraphBullets, sel, context),
        )
      } else if ('deleteParagraphBullets' in command) {
        results.push(runDeleteParagraphBullets(tr, schema, command.deleteParagraphBullets, sel))
      } else if ('updateImageProperties' in command) {
        results.push(runUpdateImageProperties(tr, command.updateImageProperties, sel))
      } else if ('insertToc' in command) {
        results.push(runInsertToc(tr, schema, command.insertToc))
      } else {
        results.push(runMoveBlocks(tr, command.moveBlocks))
      }
    }
  } catch (e) {
    return {
      ok: false,
      results: [],
      summary: '',
      error: e instanceof Error ? e.message : String(e),
    }
  }

  if (tr.steps.length > 0) {
    // route through the TrackChanges recorder so AI format/text edits become
    // reviewable revisions under the AI author (same pipeline as manual edits)
    const storage = editor.storage.trackChanges as { enabled: boolean; author: string } | undefined
    if (context.track && storage) {
      const prev = { enabled: storage.enabled, author: storage.author }
      storage.enabled = true
      storage.author = context.track.author
      try {
        editor.view.dispatch(tr)
      } finally {
        storage.enabled = prev.enabled
        storage.author = prev.author
      }
    } else {
      editor.view.dispatch(tr)
    }
  }
  return { ok: true, results, summary: summarize(results) }
}

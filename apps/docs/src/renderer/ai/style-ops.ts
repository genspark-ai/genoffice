import type { Node as PmDocNode } from '@tiptap/pm/model'
import type { StyleParaProps, StyleRunProps, StyleUpsert } from '@genoffice/docx-engine'
import type { AgentToolDef } from '../../shared/ipc'
import { parsePoints } from './lengths'
import type { Op, OpContext, OpDef, OpResult, RunEnv, SelRange, Target, TopBlock } from './ops'

/** A style as the AI sees it: the styles.xml entry plus definitions pending in this session. */
export interface AiStyleInfo {
  styleId: string
  name: string
  type: 'paragraph' | 'character' | 'table'
  basedOn?: string
  headingLevel?: number
  /** defined or changed in this session, not saved yet */
  pending?: boolean
}

/** App- or CLI-owned style catalog; upserts land in styles.xml on save. */
export interface AiStyleAccess {
  list(): AiStyleInfo[]
  /** returns an error message, or null on success */
  upsert(up: StyleUpsert): string | null
}

const STYLE_TYPES = ['paragraph', 'character', 'table'] as const
const ALIGNS = ['left', 'center', 'right', 'justify'] as const
const HIGHLIGHTS = [
  'yellow',
  'green',
  'cyan',
  'magenta',
  'blue',
  'red',
  'darkBlue',
  'darkCyan',
  'darkGreen',
  'darkMagenta',
  'darkRed',
  'darkYellow',
  'darkGray',
  'lightGray',
  'black',
  'white',
]

const LENGTH_DESC =
  'a length like "0.5in", "1.27cm", "12pt" (bare numbers are points), or null to clear'

export const DEFINE_STYLE_TOOL: AgentToolDef = {
  name: 'define_style',
  description:
    'Create a style or change the definition of an existing one (paragraph, character or table style in styles.xml). Only the fields given change; everything else the style already defines is kept. ' +
    'Paragraphs using the style pick the change up automatically. Use apply_style (or the applyStyle op) to put a style on paragraphs. Existing style ids are listed by list_styles / in the tool output.',
  inputSchema: {
    type: 'object',
    properties: {
      styleId: {
        type: 'string',
        description: 'style id (letters, digits, - and _; e.g. "Heading1", "Quote", "MyCallout")',
      },
      name: { type: 'string', description: 'display name (defaults to the id when creating)' },
      type: {
        type: 'string',
        enum: [...STYLE_TYPES],
        description: 'when creating (default paragraph); an existing style keeps its type',
      },
      basedOn: {
        type: ['string', 'null'],
        description: 'parent style id to inherit from; null removes the inheritance',
      },
      next: {
        type: ['string', 'null'],
        description:
          'style of the paragraph created by Enter at the end of a paragraph in this style',
      },
      quickFormat: { type: 'boolean', description: 'show in the quick style gallery' },
      paragraph: {
        type: 'object',
        description: 'paragraph formatting (paragraph and table styles)',
        properties: {
          align: { type: ['string', 'null'], enum: [...ALIGNS, null] },
          spaceBefore: {
            type: ['number', 'string', 'null'],
            description: `space before, ${LENGTH_DESC}`,
          },
          spaceAfter: {
            type: ['number', 'string', 'null'],
            description: `space after, ${LENGTH_DESC}`,
          },
          lineSpacing: {
            type: ['number', 'null'],
            description: 'line spacing as a multiple of single (1, 1.15, 1.5, 2)',
          },
          indentLeft: { type: ['number', 'string', 'null'], description: LENGTH_DESC },
          indentRight: { type: ['number', 'string', 'null'], description: LENGTH_DESC },
          firstLineIndent: {
            type: ['number', 'string', 'null'],
            description: `first-line indent, negative = hanging indent, ${LENGTH_DESC}`,
          },
          keepNext: { type: 'boolean', description: 'keep with next paragraph' },
          keepLines: { type: 'boolean', description: 'keep lines together' },
          pageBreakBefore: { type: 'boolean' },
          outlineLevel: {
            type: ['integer', 'null'],
            description:
              '1-9: the style becomes a heading of that level in the navigation pane / TOC',
          },
        },
      },
      run: {
        type: 'object',
        description: 'text formatting (all style types)',
        properties: {
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          underline: { type: 'boolean' },
          strike: { type: 'boolean' },
          caps: { type: 'boolean' },
          smallCaps: { type: 'boolean' },
          color: { type: ['string', 'null'], description: '"#RRGGBB" or null to clear' },
          highlight: {
            type: ['string', 'null'],
            description: `highlight color name (${HIGHLIGHTS.slice(0, 6).join(', ')}, ...) or null`,
          },
          fontSize: { type: ['number', 'null'], description: 'points' },
          fontFamily: { type: ['string', 'null'], description: 'Latin font face' },
          eastAsiaFontFamily: { type: ['string', 'null'], description: 'CJK font face' },
        },
      },
    },
    required: ['styleId'],
  },
}

export const LIST_STYLES_TOOL: AgentToolDef = {
  name: 'list_styles',
  description:
    'List the styles defined in the document (id, name, type, parent, heading level) for define_style / applyStyle.',
  inputSchema: { type: 'object', properties: {}, required: [] },
}

const HEX = /^#?([0-9a-f]{6})$/i
const STYLE_ID = /^[A-Za-z0-9_-]{1,64}$/

type Result<T> = { value: T } | { error: string }
const err = <T>(error: string): Result<T> => ({ error })

function hexOf(value: unknown, field: string): Result<string | null> {
  if (value === null) return { value: null }
  const m = typeof value === 'string' ? HEX.exec(value.trim()) : null
  return m ? { value: m[1]!.toUpperCase() } : err(`${field} must be "#RRGGBB" or null`)
}

function boolOf(value: unknown, field: string): Result<boolean | undefined> {
  if (value === undefined) return { value: undefined }
  return typeof value === 'boolean' ? { value } : err(`${field} must be true or false`)
}

function twipsOf(value: unknown, field: string): Result<number | null | undefined> {
  if (value === undefined) return { value: undefined }
  if (value === null) return { value: null }
  const pt = parsePoints(value)
  if (pt === undefined) return err(`${field} must be ${LENGTH_DESC}`)
  return { value: Math.round(pt * 20) }
}

function runProps(input: Record<string, unknown>): Result<StyleRunProps> {
  const out: StyleRunProps = {}
  for (const k of ['bold', 'italic', 'underline', 'strike', 'caps', 'smallCaps'] as const) {
    const r = boolOf(input[k], `run.${k}`)
    if ('error' in r) return r
    if (r.value !== undefined) out[k] = r.value
  }
  if (input.color !== undefined) {
    const r = hexOf(input.color, 'run.color')
    if ('error' in r) return r
    out.color = r.value
  }
  if (input.highlight !== undefined) {
    if (input.highlight !== null && !HIGHLIGHTS.includes(String(input.highlight)))
      return err(`run.highlight must be one of ${HIGHLIGHTS.join(', ')} or null`)
    out.highlight = input.highlight === null ? null : String(input.highlight)
  }
  if (input.fontSize !== undefined) {
    if (input.fontSize === null) out.sizeHalfPoints = null
    else {
      const pt = parsePoints(input.fontSize)
      if (pt === undefined || pt <= 0)
        return err('run.fontSize must be a positive number of points')
      out.sizeHalfPoints = Math.round(pt * 2)
    }
  }
  for (const [k, target] of [
    ['fontFamily', 'font'],
    ['eastAsiaFontFamily', 'eastAsiaFont'],
  ] as const) {
    const v = input[k]
    if (v === undefined) continue
    if (v !== null && (typeof v !== 'string' || !v.trim()))
      return err(`run.${k} must be a font name or null`)
    out[target] = v === null ? null : String(v).trim()
  }
  return { value: out }
}

function paraProps(input: Record<string, unknown>): Result<StyleParaProps> {
  const out: StyleParaProps = {}
  if (input.align !== undefined) {
    if (input.align !== null && !ALIGNS.includes(input.align as (typeof ALIGNS)[number]))
      return err(`paragraph.align must be one of ${ALIGNS.join(', ')} or null`)
    out.align = input.align as StyleParaProps['align']
  }
  for (const [k, target] of [
    ['spaceBefore', 'spaceBeforeTwips'],
    ['spaceAfter', 'spaceAfterTwips'],
    ['indentLeft', 'indentLeftTwips'],
    ['indentRight', 'indentRightTwips'],
    ['firstLineIndent', 'firstLineTwips'],
  ] as const) {
    const r = twipsOf(input[k], `paragraph.${k}`)
    if ('error' in r) return r
    if (r.value !== undefined) out[target] = r.value
  }
  if (input.lineSpacing !== undefined) {
    if (
      input.lineSpacing !== null &&
      !(typeof input.lineSpacing === 'number' && input.lineSpacing > 0)
    )
      return err('paragraph.lineSpacing must be a positive multiple (e.g. 1.5) or null')
    out.lineSpacing = input.lineSpacing as number | null
  }
  for (const k of ['keepNext', 'keepLines', 'pageBreakBefore'] as const) {
    const r = boolOf(input[k], `paragraph.${k}`)
    if ('error' in r) return r
    if (r.value !== undefined) out[k] = r.value
  }
  if (input.outlineLevel !== undefined) {
    const v = input.outlineLevel
    if (v !== null && !(Number.isInteger(v) && Number(v) >= 1 && Number(v) <= 9))
      return err('paragraph.outlineLevel must be 1-9 or null')
    out.outlineLevel = v as number | null
  }
  return { value: out }
}

/** Validate a define_style call against the current catalog; returns the engine upsert. */
export function resolveStyleDefinition(
  input: Record<string, unknown>,
  styles: AiStyleInfo[],
): Result<{ upsert: StyleUpsert; existing: AiStyleInfo | undefined }> {
  const styleId = typeof input.styleId === 'string' ? input.styleId.trim() : ''
  if (!STYLE_ID.test(styleId)) return err('styleId must be 1-64 letters, digits, "-" or "_"')
  const existing = styles.find((s) => s.styleId === styleId)
  const up: StyleUpsert = { styleId }
  if (input.type !== undefined) {
    if (!STYLE_TYPES.includes(input.type as (typeof STYLE_TYPES)[number]))
      return err(`type must be one of ${STYLE_TYPES.join(', ')}`)
    if (existing && existing.type !== input.type)
      return err(`${styleId} is a ${existing.type} style; its type cannot change`)
    up.type = input.type as StyleUpsert['type']
  }
  const type = existing?.type ?? up.type ?? 'paragraph'
  if (input.name !== undefined) {
    if (typeof input.name !== 'string' || !input.name.trim())
      return err('name must be a non-empty string')
    up.name = input.name.trim()
  }
  for (const k of ['basedOn', 'next'] as const) {
    const v = input[k]
    if (v === undefined) continue
    if (v === null) {
      up[k] = null
      continue
    }
    if (typeof v !== 'string') return err(`${k} must be a style id or null`)
    const parent = styles.find((s) => s.styleId === v)
    if (!parent)
      return err(
        `${k} "${v}" is not a style in this document; available ${type} styles: ${styleIds(styles, type)}`,
      )
    if (k === 'basedOn' && parent.type !== type) return err(`basedOn must be another ${type} style`)
    if (k === 'basedOn' && v === styleId) return err('a style cannot be based on itself')
    up[k] = v
  }
  const quick = boolOf(input.quickFormat, 'quickFormat')
  if ('error' in quick) return quick
  if (quick.value !== undefined) up.quickFormat = quick.value
  if (input.paragraph !== undefined) {
    if (type === 'character') return err('a character style has no paragraph formatting')
    if (!input.paragraph || typeof input.paragraph !== 'object')
      return err('paragraph must be an object')
    const r = paraProps(input.paragraph as Record<string, unknown>)
    if ('error' in r) return r
    if (Object.keys(r.value).length > 0) up.pPr = r.value
  }
  if (input.run !== undefined) {
    if (!input.run || typeof input.run !== 'object') return err('run must be an object')
    const r = runProps(input.run as Record<string, unknown>)
    if ('error' in r) return r
    if (Object.keys(r.value).length > 0) up.rPr = r.value
  }
  if (!existing && !up.pPr && !up.rPr && up.basedOn === undefined)
    return err('a new style needs paragraph or run formatting, or a basedOn parent')
  return { value: { upsert: up, existing } }
}

export function styleIds(styles: AiStyleInfo[], type?: AiStyleInfo['type']): string {
  const ids = styles.filter((s) => !type || s.type === type).map((s) => s.styleId)
  return ids.length > 40
    ? `${ids.slice(0, 40).join(', ')}, ... (${ids.length} total)`
    : ids.join(', ')
}

export function describeStyles(styles: AiStyleInfo[]): string {
  return styles
    .map((s) => {
      const bits: string[] = [s.type]
      if (s.headingLevel) bits.push(`heading ${s.headingLevel}`)
      if (s.basedOn) bits.push(`based on ${s.basedOn}`)
      if (s.pending) bits.push('unsaved')
      return `${s.styleId}${s.name !== s.styleId ? ` (${s.name})` : ''}: ${bits.join(', ')}`
    })
    .join('\n')
}

/** styleId, or the display name (case-insensitive) as a convenience */
export function findStyle(styles: AiStyleInfo[], key: string): AiStyleInfo | undefined {
  return (
    styles.find((s) => s.styleId === key) ??
    styles.find((s) => s.name.toLowerCase() === key.trim().toLowerCase())
  )
}

interface OpHelpers {
  validateShape(op: Op, def: OpDef, where: string): string | null
  matchTarget(doc: PmDocNode, target: Target, sel: SelRange): TopBlock[]
  changedAttrs(ctx: OpContext, attrs: Record<string, unknown>): Record<string, unknown>
  paragraphsIn(b: TopBlock, scoped: SelRange | null): Array<{ node: PmDocNode; pos: number }>
  scopedRange(target: Target, sel: SelRange): SelRange | null
}

/** Block ops registered by ops.ts: applyStyle puts a paragraph style on the targeted paragraphs. */
export function styleOpDefs(h: OpHelpers): OpDef[] {
  const applyStyle: OpDef = {
    name: 'applyStyle',
    signature:
      '{ op: "applyStyle", target, styleId }  // paragraph style by id (or name) from styles.xml / define_style; a heading style turns the block into a heading of its level, another paragraph style turns a heading back into a paragraph; list items keep their numbering',
    keys: ['styleId'],
    target: 'required',
    validate(op, where) {
      const shape = h.validateShape(op, this, where)
      if (shape) return shape
      if (typeof op.styleId !== 'string' || !op.styleId.trim())
        return `${where}: styleId must be a non-empty string`
      return null
    },
    apply(op, env: RunEnv): OpResult {
      const { tr, schema, ctx, sel } = env
      const catalog = ctx.styles?.list()
      if (!catalog) throw new Error('applyStyle: the style catalog is not available here')
      const style = findStyle(catalog, String(op.styleId))
      if (!style)
        throw new Error(
          `applyStyle: no style "${String(op.styleId)}" in this document; paragraph styles: ${styleIds(catalog, 'paragraph')}. Create one with define_style first.`,
        )
      if (style.type !== 'paragraph')
        throw new Error(
          `applyStyle: ${style.styleId} is a ${style.type} style; only paragraph styles apply to blocks`,
        )
      const target = op.target as Target
      const matched = h.matchTarget(tr.doc, target, sel)
      const scoped = h.scopedRange(target, sel)
      let changed = 0
      let skippedProtected = 0
      for (const b of matched) {
        if (b.node.type.name === 'docProtected') {
          skippedProtected++
          continue
        }
        for (const p of h.paragraphsIn(b, scoped)) {
          const isHeading = p.node.type.name === 'docHeading'
          const level = style.headingLevel
          let type = level
            ? schema.nodes.docHeading
            : isHeading
              ? schema.nodes.docParagraph
              : p.node.type
          // table cells take no docHeading: the paragraph keeps its type and carries the style by id
          const $p = tr.doc.resolve(p.pos)
          if (type !== p.node.type && !$p.parent.canReplaceWith($p.index(), $p.index() + 1, type))
            type = p.node.type
          const sameType = type === p.node.type
          const asHeading = level && type === schema.nodes.docHeading
          if (
            sameType &&
            p.node.attrs.styleId === style.styleId &&
            (!asHeading || Number(p.node.attrs.level) === level)
          )
            continue
          const attrs: Record<string, unknown> = { ...p.node.attrs, styleId: style.styleId }
          if (asHeading) attrs.level = level
          tr.setNodeMarkup(p.pos, sameType ? undefined : type, h.changedAttrs(ctx, attrs))
          changed++
        }
      }
      return { op: 'applyStyle', matched: matched.length, changed, skippedProtected }
    },
  }
  return [applyStyle]
}

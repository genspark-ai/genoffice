import type { Editor, JSONContent } from '@tiptap/core'
import {
  buildAnchoredTextboxParagraphXml,
  type ImageWrap,
  type NewImage,
  type PictureWatermarkSpec,
  type Run,
  type TextboxDisplay,
  type Watermark,
  type WatermarkSpec,
} from '@genoffice/docx-engine'
import type { AgentToolDef } from '../../shared/ipc'
import { blockRangePositions } from './doc-utils'
import { emuToPx, parseEmu, parsePoints } from './lengths'

/** App- or CLI-owned watermark state; the header part is rewritten on save. */
export interface AiWatermarkAccess {
  /** current text watermark, null when the header has none (or carries a picture) */
  current(): string | null
  /** returns an error message, or null on success */
  set(spec: WatermarkSpec | PictureWatermarkSpec | null): string | null
}

export type { Watermark }

const HEX = /^#?([0-9a-f]{6})$/i
const LENGTH_DESC = '"2.54cm", "1in", "72pt", "96px" or a bare number of points'

export const SET_WATERMARK_TOOL: AgentToolDef = {
  name: 'set_watermark',
  description:
    'Set, restyle or remove the page watermark: a gray text like DRAFT or CONFIDENTIAL, or a washed-out picture (a logo), drawn behind the body on every page (written into the page header, where Word shows it in Design > Watermark). Give text or image, not both; text null or image null removes whichever watermark the document has.',
  inputSchema: {
    type: 'object',
    properties: {
      text: { type: ['string', 'null'], description: 'watermark text; null removes the watermark' },
      fontFamily: { type: 'string', description: 'font face (default the document theme face)' },
      color: { type: 'string', description: '"#RRGGBB" (default silver)' },
      opacity: { type: 'number', description: '0-1 (default 0.5)' },
      diagonal: { type: 'boolean', description: 'diagonal (default) or horizontal layout' },
      bold: { type: 'boolean' },
      italic: { type: 'boolean' },
      image: {
        type: ['string', 'null'],
        description:
          'picture watermark: http(s) or data: URL of a png/jpg/gif, centered behind the body; null removes the watermark',
      },
      scale: {
        type: 'number',
        description:
          'picture size as a percent of its natural size (default: fit the page inside the margins)',
      },
      washout: {
        type: 'boolean',
        description: "fade the picture like Word's washout (default true)",
      },
    },
  },
}

export interface ResolvedWatermarkImage {
  base64: string
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
  widthPx: number
  heightPx: number
}

/** Field check for the picture variant; the caller has already fetched and measured the image. */
export function resolvePictureWatermark(
  input: Record<string, unknown>,
  image: ResolvedWatermarkImage,
): { spec: PictureWatermarkSpec } | { error: string } {
  if (input.text !== undefined && input.text !== null)
    return { error: 'give text or image, not both' }
  const spec: PictureWatermarkSpec = { image }
  if (input.scale !== undefined) {
    if (typeof input.scale !== 'number' || !(input.scale > 0) || input.scale > 1000)
      return { error: 'scale must be a percent between 0 and 1000' }
    spec.scale = input.scale
  }
  if (input.washout !== undefined) {
    if (typeof input.washout !== 'boolean') return { error: 'washout must be true or false' }
    spec.washout = input.washout
  }
  return { spec }
}

export function resolveWatermark(
  input: Record<string, unknown>,
): { spec: WatermarkSpec | null } | { error: string } {
  if (input.text === null || (input.image === null && input.text === undefined))
    return { spec: null }
  if (typeof input.image === 'string')
    return { error: 'a picture watermark needs the image resolved first' }
  if (typeof input.text !== 'string' || !input.text.trim())
    return {
      error:
        'text must be a non-empty string (or image an image URL); null for either removes the watermark',
    }
  if (input.text.length > 200)
    return { error: 'text is too long for a watermark (200 characters max)' }
  const spec: WatermarkSpec = { text: input.text.trim() }
  if (input.fontFamily !== undefined) {
    if (typeof input.fontFamily !== 'string' || !input.fontFamily.trim())
      return { error: 'fontFamily must be a font name' }
    spec.fontFamily = input.fontFamily.trim()
  }
  if (input.color !== undefined) {
    const m = typeof input.color === 'string' ? HEX.exec(input.color.trim()) : null
    if (!m) return { error: 'color must be "#RRGGBB"' }
    spec.colorHex = m[1]!.toUpperCase()
  }
  if (input.opacity !== undefined) {
    if (typeof input.opacity !== 'number' || input.opacity < 0 || input.opacity > 1)
      return { error: 'opacity must be between 0 and 1' }
    spec.opacity = input.opacity
  }
  for (const k of ['diagonal', 'bold', 'italic'] as const) {
    if (input[k] === undefined) continue
    if (typeof input[k] !== 'boolean') return { error: `${k} must be true or false` }
    spec[k] = input[k] as boolean
  }
  return { spec }
}

const ANCHORS = ['paragraph', 'page', 'margin'] as const
type Anchor = (typeof ANCHORS)[number]
const WRAPS = ['square', 'tight', 'topAndBottom', 'behind', 'inFront'] as const
type Wrap = (typeof WRAPS)[number]

const WRAP_TO_IMAGE: Record<Wrap, ImageWrap> = {
  square: 'square-left',
  tight: 'tight-left',
  topAndBottom: 'topBottom',
  behind: 'behind',
  inFront: 'front',
}

export interface FloatSpec {
  anchor: Anchor
  xEmu: number
  yEmu: number
  wrap: Wrap
}

const FLOAT_SCHEMA = {
  type: 'object',
  description:
    'omit for an inline picture in its own paragraph; give it to float the picture at a position',
  properties: {
    anchor: {
      type: 'string',
      enum: [...ANCHORS],
      description:
        'what x/y measure from: the anchor paragraph (x from the text column, y from the paragraph top), the page corner, or the margin box corner',
    },
    x: { type: ['number', 'string'], description: `horizontal offset, ${LENGTH_DESC}` },
    y: { type: ['number', 'string'], description: `vertical offset, ${LENGTH_DESC}` },
    wrap: {
      type: 'string',
      enum: [...WRAPS],
      description: 'how body text flows around it (default square)',
    },
  },
  required: ['x', 'y'],
}

export function resolveFloat(value: unknown): { float: FloatSpec } | { error: string } {
  if (!value || typeof value !== 'object')
    return { error: 'float must be an object { anchor?, x, y, wrap? }' }
  const f = value as Record<string, unknown>
  const anchor = f.anchor === undefined ? 'paragraph' : f.anchor
  if (!ANCHORS.includes(anchor as Anchor))
    return { error: `float.anchor must be one of ${ANCHORS.join(', ')}` }
  const wrap = f.wrap === undefined ? 'square' : f.wrap
  if (!WRAPS.includes(wrap as Wrap))
    return { error: `float.wrap must be one of ${WRAPS.join(', ')}` }
  const xEmu = parseEmu(f.x, 'pt')
  const yEmu = parseEmu(f.y, 'pt')
  if (xEmu === undefined || yEmu === undefined)
    return { error: `float.x and float.y must be ${LENGTH_DESC}` }
  return { float: { anchor: anchor as Anchor, xEmu, yEmu, wrap: wrap as Wrap } }
}

export const INSERT_PICTURE_TOOL: AgentToolDef = {
  name: 'insert_picture',
  description:
    'Insert a picture from a direct image link at a block position, at a given size, inline or floating (positioned on the page, with text wrapping around/behind/in front). insert_image is the simpler cursor-position variant.',
  inputSchema: {
    type: 'object',
    properties: {
      url: { type: 'string', description: 'direct image link (png/jpg/gif)' },
      afterBlockIndex: {
        type: 'integer',
        description: 'insert after this block (-1 = start of document); default: end of document',
      },
      width: {
        type: ['number', 'string'],
        description: `display width, ${LENGTH_DESC}; default natural size capped at 480px`,
      },
      height: {
        type: ['number', 'string'],
        description: 'display height; omit to keep the aspect ratio',
      },
      float: FLOAT_SCHEMA,
      altText: { type: 'string', description: 'alternative text (accessibility)' },
    },
    required: ['url'],
  },
}

export interface PictureInput {
  base64: string
  mime: NewImage['mime']
  naturalWidth: number
  naturalHeight: number
  width?: unknown
  height?: unknown
  float?: FloatSpec
  altText?: string
  label?: string
}

/** The protected image block for a new picture: inline, or anchored when `float` is given. */
export function pictureNode(
  input: PictureInput,
): { node: JSONContent; widthPx: number; heightPx: number } | { error: string } {
  const wEmu = input.width === undefined ? undefined : parseEmu(input.width, 'pt')
  const hEmu = input.height === undefined ? undefined : parseEmu(input.height, 'pt')
  if ((input.width !== undefined && !wEmu) || (input.height !== undefined && !hEmu))
    return { error: `width and height must be ${LENGTH_DESC}` }
  const ratio = input.naturalWidth / Math.max(1, input.naturalHeight)
  let widthPx: number
  let heightPx: number
  if (wEmu && hEmu) {
    widthPx = emuToPx(wEmu)
    heightPx = emuToPx(hEmu)
  } else if (wEmu) {
    widthPx = emuToPx(wEmu)
    heightPx = Math.round(widthPx / ratio)
  } else if (hEmu) {
    heightPx = emuToPx(hEmu)
    widthPx = Math.round(heightPx * ratio)
  } else {
    const scale = Math.min(1, 480 / input.naturalWidth)
    widthPx = Math.round(input.naturalWidth * scale)
    heightPx = Math.round(input.naturalHeight * scale)
  }
  widthPx = Math.max(1, widthPx)
  heightPx = Math.max(1, heightPx)
  const genImage: NewImage = {
    base64: input.base64,
    mime: input.mime,
    widthPx,
    heightPx,
    ...(input.altText ? { altText: input.altText } : {}),
  }
  const attrs: Record<string, unknown> = {
    docxIndex: null,
    blockType: 'image',
    label: input.label ?? 'Picture',
    imageDataUrl: `data:${input.mime};base64,${input.base64}`,
    imageWidthPx: widthPx,
    imageHeightPx: heightPx,
  }
  if (input.float) {
    const { anchor, xEmu, yEmu, wrap } = input.float
    genImage.wrap = WRAP_TO_IMAGE[wrap]
    genImage.posOffsetEmu = {
      x: xEmu,
      y: yEmu,
      ...(anchor === 'paragraph' ? {} : { relativeTo: anchor }),
    }
    attrs.imageWrap = genImage.wrap
    attrs.imageOffsetXEmu = xEmu
    attrs.imageOffsetYEmu = yEmu
    if (anchor !== 'paragraph') attrs.imageRelV = anchor
  }
  attrs.genImage = genImage
  return { node: { type: 'docProtected', attrs }, widthPx, heightPx }
}

/** top-level position after block `after` (-1 = document start); error text when out of range */
export function insertPosition(
  editor: Editor,
  after: unknown,
): { pos: number; after: number } | { error: string } {
  const count = editor.state.doc.childCount
  const index = after === undefined || after === null ? count - 1 : Number(after)
  if (!Number.isInteger(index) || index < -1 || index >= count)
    return { error: `afterBlockIndex must be -1..${count - 1} (the document has ${count} blocks)` }
  return { pos: index < 0 ? 0 : blockRangePositions(editor, index, index).to, after: index }
}

export const INSERT_TEXT_BOX_TOOL: AgentToolDef = {
  name: 'insert_text_box',
  description:
    'Insert a floating text box (a positioned rectangle holding text, like a callout or sidebar) anchored to a block. Body text either flows above/below it or ignores it.',
  inputSchema: {
    type: 'object',
    properties: {
      afterBlockIndex: {
        type: 'integer',
        description: 'anchor after this block (-1 = start of document); default: end of document',
      },
      text: { type: 'string', description: 'box text; \\n separates paragraphs' },
      width: { type: ['number', 'string'], description: `box width, ${LENGTH_DESC}` },
      height: { type: ['number', 'string'], description: `box height, ${LENGTH_DESC}` },
      x: { type: ['number', 'string'], description: `horizontal offset, ${LENGTH_DESC}` },
      y: { type: ['number', 'string'], description: `vertical offset, ${LENGTH_DESC}` },
      anchor: {
        type: 'string',
        enum: ['paragraph', 'page'],
        description: 'offsets measure from the anchor paragraph (default) or the page corner',
      },
      wrap: {
        type: 'string',
        enum: ['topAndBottom', 'none'],
        description:
          'body text resumes below the box (default for paragraph anchors) or ignores it (default for page anchors)',
      },
      fill: {
        type: ['string', 'null'],
        description: '"#RRGGBB" background (default white); null = transparent',
      },
      borderColor: {
        type: ['string', 'null'],
        description: '"#RRGGBB" outline (default black); null = no outline',
      },
      fontSize: { type: 'number', description: 'text size in points' },
      align: { type: 'string', enum: ['left', 'center', 'right', 'justify'] },
      bold: { type: 'boolean' },
      color: { type: 'string', description: '"#RRGGBB" text color' },
    },
    required: ['text', 'width', 'height', 'x', 'y'],
  },
}

function hexOrNull(
  value: unknown,
  field: string,
  dflt: string,
): { hex: string | undefined } | { error: string } {
  if (value === undefined) return { hex: dflt }
  if (value === null) return { hex: undefined }
  const m = typeof value === 'string' ? HEX.exec(value.trim()) : null
  return m ? { hex: m[1]!.toUpperCase() } : { error: `${field} must be "#RRGGBB" or null` }
}

/** Validate an insert_text_box call and build the anchored paragraph + its display model. */
export function textBoxNode(
  input: Record<string, unknown>,
): { node: JSONContent; widthPx: number; heightPx: number } | { error: string } {
  const text = typeof input.text === 'string' ? input.text : ''
  if (!text.trim()) return { error: 'text must not be empty' }
  const widthEmu = parseEmu(input.width, 'pt')
  const heightEmu = parseEmu(input.height, 'pt')
  const xEmu = parseEmu(input.x, 'pt')
  const yEmu = parseEmu(input.y, 'pt')
  if (!widthEmu || !heightEmu || widthEmu <= 0 || heightEmu <= 0)
    return { error: `width and height must be positive lengths: ${LENGTH_DESC}` }
  if (xEmu === undefined || yEmu === undefined) return { error: `x and y must be ${LENGTH_DESC}` }
  const anchor = input.anchor === undefined ? 'paragraph' : input.anchor
  if (anchor !== 'paragraph' && anchor !== 'page')
    return { error: 'anchor must be "paragraph" or "page"' }
  const wrap = input.wrap === undefined ? (anchor === 'page' ? 'none' : 'topAndBottom') : input.wrap
  if (wrap !== 'topAndBottom' && wrap !== 'none')
    return { error: 'wrap must be "topAndBottom" or "none"' }
  const fill = hexOrNull(input.fill, 'fill', 'FFFFFF')
  if ('error' in fill) return fill
  const border = hexOrNull(input.borderColor, 'borderColor', '000000')
  if ('error' in border) return border
  const run: Run = { text: '' }
  if (input.fontSize !== undefined) {
    const pt = parsePoints(input.fontSize)
    if (pt === undefined || pt <= 0)
      return { error: 'fontSize must be a positive number of points' }
    run.sizeHalfPoints = Math.round(pt * 2)
  }
  if (input.bold === true) run.bold = true
  if (input.color !== undefined) {
    const m = typeof input.color === 'string' ? HEX.exec(input.color.trim()) : null
    if (!m) return { error: 'color must be "#RRGGBB"' }
    run.color = m[1]!.toUpperCase()
  }
  const align = input.align as 'left' | 'center' | 'right' | 'justify' | undefined
  if (align !== undefined && !['left', 'center', 'right', 'justify'].includes(align))
    return { error: 'align must be left, center, right or justify' }
  const paragraphs = text.split('\n').map((line) => ({
    runs: [{ ...run, text: line }],
    ...(align ? { format: { align } } : {}),
  }))
  const xml = buildAnchoredTextboxParagraphXml({
    anchor,
    xEmu,
    yEmu,
    widthEmu,
    heightEmu,
    paragraphs,
    id: Math.floor(Math.random() * 900000) + 100000,
    fillHex: fill.hex,
    borderHex: border.hex,
    wrap,
  })
  const widthPx = emuToPx(widthEmu)
  const heightPx = emuToPx(heightEmu)
  const display: TextboxDisplay = {
    ...(fill.hex ? { fill: fill.hex } : {}),
    ...(border.hex ? { borderColor: border.hex } : {}),
    widthPx,
    heightPx,
    paras: paragraphs.map((p) => ({ runs: p.runs, ...(align ? { align } : {}) })),
    offsetXEmu: xEmu,
    offsetYEmu: yEmu,
    ...(wrap === 'none'
      ? { floating: true }
      : { bandTopPx: emuToPx(yEmu), bandBottomPx: emuToPx(yEmu) + heightPx }),
    ...(anchor === 'page' ? { pagePinned: true, floating: true } : {}),
  }
  return {
    node: {
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'passthrough',
        label: 'Text box',
        previewText: text.replace(/\s+/g, ' ').trim(),
        genXml: xml,
        textboxes: [display],
      },
    },
    widthPx,
    heightPx,
  }
}

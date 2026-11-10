import type { PDFPageProxy } from 'pdfjs-dist'
import { OPS } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { FormFieldInput, FormFieldKind } from '../shared/ipc'

type Rect = [number, number, number, number]

/** A positioned text run in PDF user space (baseline origin, y up) */
export interface DetectTextItem {
  str: string
  x: number
  y: number
  w: number
  h: number
}

/** Bounding box of one drawn path in PDF user space */
export type DetectShape = Rect

export interface DetectPageInput {
  pageIndex: number
  /** CropBox in PDF user space */
  pageBox: Rect
  items: DetectTextItem[]
  shapes: DetectShape[]
  /** Rects already holding a widget or pending field; candidates there are dropped */
  occupied: Rect[]
  /** Names already in use; the detector appends unique suffixes */
  taken: ReadonlySet<string>
}

const MIN_FIELD_H = 16
const SIGNATURE_H = 36
const MIN_GAP_W = 40
const UNDERLINE_MAX_H = 1.6
const CHECK_GLYPHS = new Set(['☐', '□', '◻', '▢', '○', '◯', '[ ]', '( )', '[]', '()'])

const area = (r: Rect) => Math.max(0, r[2] - r[0]) * Math.max(0, r[3] - r[1])
const intersection = (a: Rect, b: Rect): Rect => [
  Math.max(a[0], b[0]),
  Math.max(a[1], b[1]),
  Math.min(a[2], b[2]),
  Math.min(a[3], b[3]),
]
const overlapRatio = (a: Rect, b: Rect) => {
  const i = area(intersection(a, b))
  return i === 0 ? 0 : i / Math.min(area(a), area(b))
}
const itemRect = (it: DetectTextItem): Rect => [it.x, it.y - it.h * 0.25, it.x + it.w, it.y + it.h]

const sameLine = (a: DetectTextItem, b: DetectTextItem) =>
  Math.abs(a.y - b.y) < Math.max(a.h, b.h) * 0.6

/** Field kind implied by the label text; plain text otherwise */
function kindForLabel(label: string): { kind: FormFieldKind; dateFormat?: string } {
  if (/\b(sign(ature)?|signed by)\b|\u7b7e\u540d|\u7c3d\u540d|\u7f72\u540d/i.test(label))
    return { kind: 'signature' }
  if (/\b(date|dob|birth(day)?)\b|\u65e5\u671f|\u5e74\u6708\u65e5|\u751f\u65e5/i.test(label))
    return { kind: 'text', dateFormat: 'yyyy-mm-dd' }
  return { kind: 'text' }
}

/** ASCII identifier from a label ("Full name:" → full_name); empty when nothing survives */
export function nameFromLabel(label: string): string {
  return label
    .normalize('NFKD')
    .replace(/[:：*＊]+\s*$/, '')
    .replace(/[^A-Za-z0-9]+/g, ' ')
    .trim()
    .toLowerCase()
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 4)
    .join('_')
}

interface Candidate {
  rect: Rect
  kind: FormFieldKind
  label: string
  dateFormat?: string
}

/**
 * Find places a reader would be expected to write: underscore runs, drawn
 * underlines and boxes without text in them, empty space after a "Label:" and
 * checkbox glyphs. Returns ready-to-add fields named after their labels.
 */
export function detectFormFields(input: DetectPageInput): FormFieldInput[] {
  const { items, shapes, pageBox } = input
  const textItems = items.filter((it) => it.str.trim().length > 0 && it.w > 0 && it.h > 0)
  const candidates: Candidate[] = []

  const textInside = (r: Rect) =>
    textItems.some((it) => {
      const ir = itemRect(it)
      return overlapRatio(ir, r) > 0.3 && /[^\s_]/.test(it.str)
    })

  /** Nearest text run that reads as this field's caption: left on the same line (right for a checkbox), else just above */
  const labelFor = (r: Rect, kind: FormFieldKind): string => {
    const cy = (r[1] + r[3]) / 2
    if (kind === 'checkbox') {
      const right = textItems
        .filter(
          (it) =>
            it.x >= r[2] - 2 &&
            it.x - r[2] < 40 &&
            Math.abs(it.y + it.h * 0.35 - cy) < Math.max(it.h, r[3] - r[1]) * 0.7 &&
            !CHECK_GLYPHS.has(it.str.trim()),
        )
        .sort((a, b) => a.x - b.x)[0]
      if (right) return right.str.trim()
    }
    const left = textItems
      .filter(
        (it) =>
          it.x + it.w <= r[0] + 2 &&
          r[0] - (it.x + it.w) < 120 &&
          Math.abs(it.y + it.h * 0.35 - cy) < Math.max(it.h, r[3] - r[1]) * 0.7 &&
          !/^[_\s]+$/.test(it.str),
      )
      .sort((a, b) => b.x + b.w - (a.x + a.w))[0]
    if (left) return left.str.trim()
    const above = textItems
      .filter(
        (it) =>
          it.y > r[3] - 1 &&
          it.y - r[3] < Math.max(it.h * 1.8, 14) &&
          it.x < r[2] &&
          it.x + it.w > r[0] - 2 &&
          !/^[_\s]+$/.test(it.str),
      )
      .sort((a, b) => a.y - b.y)[0]
    return above ? above.str.trim() : ''
  }

  // Clamp before the overlap pass so a box pulled back onto the page still competes fairly
  const clamp = (r: Rect): Rect => {
    const w = Math.min(r[2] - r[0], pageBox[2] - pageBox[0])
    const h = Math.min(r[3] - r[1], pageBox[3] - pageBox[1])
    const x = Math.min(Math.max(r[0], pageBox[0]), pageBox[2] - w)
    const y = Math.min(Math.max(r[1], pageBox[1]), pageBox[3] - h)
    return [x, y, x + w, y + h]
  }

  const push = (rect: Rect, kind: FormFieldKind, label = labelFor(rect, kind)) => {
    const hinted = kind === 'text' ? kindForLabel(label) : { kind }
    let r = rect
    if (hinted.kind === 'signature' && r[3] - r[1] < SIGNATURE_H)
      r = [r[0], r[3] - SIGNATURE_H, r[2], r[3]]
    r = clamp(r)
    candidates.push({
      rect: r,
      kind: hinted.kind,
      label,
      ...('dateFormat' in hinted ? { dateFormat: hinted.dateFormat } : {}),
    })
  }

  // 1. Underscore runs inside text: "Name: ________"
  for (const it of textItems) {
    const re = /_{4,}/g
    let m: RegExpExecArray | null
    while ((m = re.exec(it.str))) {
      const cw = it.w / it.str.length
      const x1 = it.x + m.index * cw
      const x2 = x1 + m[0].length * cw
      if (x2 - x1 < MIN_GAP_W * 0.6) continue
      const h = Math.max(MIN_FIELD_H, it.h * 1.3)
      const before = it.str.slice(0, m.index).trim()
      const rect: Rect = [x1, it.y - 2, x2, it.y - 2 + h]
      push(rect, 'text', before || labelFor(rect, 'text'))
    }
  }

  // 2. Checkbox glyphs, alone or leading a caption ("[ ] I agree")
  for (const it of textItems) {
    const m = /^\s*(\[\s?\]|\(\s?\)|[☐□◻▢○◯])(?:\s+(.*))?$/.exec(it.str)
    if (!m) continue
    const cw = it.w / it.str.length
    const glyphW = m[1]!.length * cw
    const size = Math.max(8, Math.min(it.h, Math.max(glyphW, it.h * 0.8)))
    const rect: Rect = [it.x, it.y, it.x + size, it.y + size]
    push(rect, 'checkbox', m[2]?.trim() || labelFor(rect, 'checkbox'))
  }

  // 3. Drawn underlines and boxes
  for (const s of shapes) {
    const w = s[2] - s[0]
    const h = s[3] - s[1]
    if (h <= UNDERLINE_MAX_H && w >= MIN_GAP_W) {
      const r: Rect = [s[0], s[1] + 1, s[2], s[1] + 1 + Math.max(MIN_FIELD_H, 18)]
      if (!textInside(r)) push(r, 'text')
      continue
    }
    if (w >= 7 && w <= 20 && Math.abs(w - h) <= 2.5) {
      if (!textInside(s)) push(s, 'checkbox')
      continue
    }
    if (w >= 30 && w <= 520 && h >= 10 && h <= 60) {
      if (!textInside(s)) push([s[0] + 1, s[1] + 1, s[2] - 1, s[3] - 1], 'text')
    }
  }

  // 4. "Label:" followed by empty space on the same line
  const marginRight = pageBox[2] - 36
  for (const it of textItems) {
    if (!/[:：]\s*$/.test(it.str)) continue
    const right = textItems
      .filter((o) => o !== it && sameLine(it, o) && o.x > it.x + it.w - 1)
      .sort((a, b) => a.x - b.x)[0]
    const x1 = it.x + it.w + 4
    const x2 = (right ? right.x : marginRight) - 4
    if (x2 - x1 < MIN_GAP_W) continue
    const h = Math.max(MIN_FIELD_H, it.h * 1.3)
    push([x1, it.y - h * 0.22, x2, it.y - h * 0.22 + h], 'text', it.str.trim())
  }

  // Keep the first candidate of each spot; drop anything on an existing widget
  const kept: Candidate[] = []
  for (const c of candidates) {
    if (c.rect[2] - c.rect[0] < 6 || c.rect[3] - c.rect[1] < 6) continue
    if (input.occupied.some((o) => overlapRatio(o, c.rect) > 0.2)) continue
    if (kept.some((k) => overlapRatio(k.rect, c.rect) > 0.3)) continue
    kept.push(c)
  }
  // Top-to-bottom, left-to-right like a reader fills the form
  kept.sort((a, b) => b.rect[3] - a.rect[3] || a.rect[0] - b.rect[0])

  const used = new Set(input.taken)
  const unique = (base: string) => {
    if (base && !used.has(base)) {
      used.add(base)
      return base
    }
    for (let n = 1; ; n++) {
      const name = `${base || 'field'}_${n}`
      if (!used.has(name)) {
        used.add(name)
        return name
      }
    }
  }
  return kept.map((c) => ({
    name: unique(nameFromLabel(c.label) || (c.kind === 'text' ? 'field' : c.kind)),
    kind: c.kind,
    pageIndex: input.pageIndex,
    rect: c.rect.map((v) => Math.round(v * 100) / 100) as Rect,
    ...(c.dateFormat ? { dateFormat: c.dateFormat } : {}),
  }))
}

type Matrix = [number, number, number, number, number, number]
const mul = (m: Matrix, n: Matrix): Matrix => [
  m[0] * n[0] + m[1] * n[2],
  m[0] * n[1] + m[1] * n[3],
  m[2] * n[0] + m[3] * n[2],
  m[2] * n[1] + m[3] * n[3],
  m[4] * n[0] + m[5] * n[2] + n[4],
  m[4] * n[1] + m[5] * n[3] + n[5],
]
const apply = (m: Matrix, x: number, y: number): [number, number] => [
  m[0] * x + m[2] * y + m[4],
  m[1] * x + m[3] * y + m[5],
]

/** Text runs and drawn-path boxes of a page in PDF user space (pdf.js operator list, CTM tracked) */
export async function collectDetectInput(
  page: PDFPageProxy,
): Promise<Pick<DetectPageInput, 'items' | 'shapes'>> {
  const content = await page.getTextContent()
  const items: DetectTextItem[] = []
  for (const raw of content.items as {
    str?: string
    transform?: number[]
    width?: number
    height?: number
  }[]) {
    if (typeof raw.str !== 'string' || !raw.transform) continue
    const h = raw.height || Math.hypot(raw.transform[2] ?? 0, raw.transform[3] ?? 0)
    if (Math.abs(raw.transform[1] ?? 0) > h * 1e-3) continue
    items.push({
      str: raw.str,
      x: raw.transform[4] ?? 0,
      y: raw.transform[5] ?? 0,
      w: raw.width ?? 0,
      h,
    })
  }

  const ops = await page.getOperatorList()
  const shapes: DetectShape[] = []
  let ctm: Matrix = [1, 0, 0, 1, 0, 0]
  const stack: Matrix[] = []
  for (let i = 0; i < ops.fnArray.length; i++) {
    const fn = ops.fnArray[i]
    const args = ops.argsArray[i] as unknown[]
    switch (fn) {
      case OPS.save:
        stack.push(ctm)
        break
      case OPS.restore:
        ctm = stack.pop() ?? ctm
        break
      case OPS.transform:
        ctm = mul(args as Matrix, ctm)
        break
      case OPS.paintFormXObjectBegin: {
        stack.push(ctm)
        const m = args[0] as Matrix | null
        if (m) ctm = mul(m, ctm)
        break
      }
      case OPS.paintFormXObjectEnd:
        ctm = stack.pop() ?? ctm
        break
      case OPS.constructPath: {
        // args = [paintOp, pathData, minMax]; minMax is the untransformed path bbox
        const minMax = args[2] as number[] | undefined
        const paint = args[0] as number
        if (!minMax || paint === OPS.endPath) break
        const pts = [
          apply(ctm, minMax[0]!, minMax[1]!),
          apply(ctm, minMax[2]!, minMax[1]!),
          apply(ctm, minMax[0]!, minMax[3]!),
          apply(ctm, minMax[2]!, minMax[3]!),
        ]
        const xs = pts.map((p) => p[0])
        const ys = pts.map((p) => p[1])
        const r: Rect = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]
        if (Number.isFinite(r[0] + r[1] + r[2] + r[3])) shapes.push(r)
        break
      }
    }
  }
  return { items, shapes }
}

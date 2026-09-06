/**
 * Slide edit script sandbox (modeled on the Google Slides add-in's execute_apps_script idea):
 * instead of blindly sending absolute coordinates / calling fine-grained tools attribute by
 * attribute, the AI submits a synchronous JS-shaped DSL program — at runtime it gets a snapshot of
 * this page's real element geometry and text (els) and the canvas size (canvas), computes
 * algorithmically, then uses setBox/moveBy/resizeBy for geometry and setText/setStyle/setFill/setStroke
 * for content and style. The script only collects operations and never mutates the screen directly;
 * the caller dispatches the collected result to existing edit IPC (geometry applied atomically in
 * one batchEditTransform, the rest serially in script order).
 *
 * All precision-number work (reading coordinates, computing spacing) happens at execution site;
 * coordinates never round-trip through LLM tokens — the key to getting layout adjustments right.
 */

import type { EditParagraph } from '../../shared/ipc'
import { interpretLayoutScript } from './layout-script-interpreter'

export interface LayoutScriptElement {
  id: string
  type: string
  /** Element text (tables joined row by row; empty string when no text) */
  text: string
  x: number
  y: number
  w: number
  h: number
  rotation: number
  /** Max font size of the text (pt; omitted when no text) */
  fontSizePt?: number
  /** Solid fill color (#RRGGBB; omitted for none/gradient/image fills) */
  fill?: string
  /** Dominant text color (#RRGGBB; omitted when no text) */
  textColor?: string
  /** Stroke/border color (#RRGGBB; omitted when no stroke) */
  strokeColor?: string
  /** Group children: coordinates are absolute; direct children of a top-level group are editable (ops carry groupId), deeper nesting is read-only */
  inGroup?: boolean
  /** Direct child of a top-level group: the parent group's id (write ops route through the in-group edit pipeline) */
  groupId?: string
  /** master/layout decoration: read-only, not modifiable */
  locked?: boolean
}

export interface LayoutOp {
  id: string
  x: number
  y: number
  w: number
  h: number
  rotation: number
  /** Group child: applied via the per-element in-group transform IPC instead of the batch */
  groupId?: string
}

/** setStyle's style-override fields (pass only what changes; align is set on the paragraph, the rest override per run) */
export interface SlideStylePatch {
  fontSize?: number
  color?: string
  bold?: boolean
  italic?: boolean
  underline?: boolean
  fontFamily?: string
  align?: 'left' | 'center' | 'right'
}

/** Non-geometry edit operations (collected in script call order; the outer layer dispatches them in order to the matching IPC) */
export type SlideEditOp = (
  | { kind: 'text'; id: string; paragraphs: EditParagraph[] }
  | { kind: 'style'; id: string; style: SlideStylePatch }
  | { kind: 'fill'; id: string; fill: string }
  | { kind: 'stroke'; id: string; stroke: { color: string; widthPt: number } | null }
) & {
  /** Group child: dispatched with groupId so the IPC patches the child slice inside the group */
  groupId?: string
}

export interface LayoutScriptResult {
  /** Geometry ops (merged result of setBox/moveBy/resizeBy, applied once via batchEditTransform) */
  ops: LayoutOp[]
  /** Text/style/fill/stroke ops (kept in script call order) */
  edits: SlideEditOp[]
  /** The script's return value (JSON-serialized and echoed back to the AI) */
  returned?: string
  /** In-script log() output (for debugging, echoed back to the AI) */
  logs: string[]
  error?: string
}

/** #RRGGBB (# optional) → normalized with # prefix; invalid returns null */
function normalizeHex(v: unknown): string | null {
  if (typeof v !== 'string') return null
  const s = v.trim()
  if (!/^#?[0-9a-fA-F]{6}$/.test(s)) return null
  return s.startsWith('#') ? s : `#${s}`
}

export function runLayoutScript(
  code: string,
  elements: LayoutScriptElement[],
  canvas: { w: number; h: number },
): LayoutScriptResult {
  const logs: string[] = []

  const byId = new Map(elements.map((e) => [e.id, e]))
  // setBox results: id → full target box (patch merged on top of the current value; repeated calls, later wins)
  const pending = new Map<string, LayoutOp>()
  // Non-geometry ops: keep script call order (order matters when the same element changes text then style)
  const edits: SlideEditOp[] = []

  /** Common existence/read-only validation (shared by all write primitives).
   *  Direct children of a top-level group pass (their ops carry groupId); deeper nesting is read-only. */
  const guard = (fn: string, id: unknown): LayoutScriptElement => {
    const key = String(id)
    const el = byId.get(key)
    if (!el) throw new Error(`${fn}: element "${key}" does not exist (see the ids in els)`)
    if (el.inGroup && !el.groupId)
      throw new Error(
        `${fn}: "${key}" is nested inside a sub-group (read-only); ungroup the outer group first (apply_ops ungroupElement), or operate on the sub-group as a whole`,
      )
    if (el.locked)
      throw new Error(`${fn}: "${key}" is a layout decoration element, read-only and unmodifiable`)
    return el
  }

  /** Required-number validation (for moveBy/resizeBy delta arguments) */
  const reqNum = (ctx: string, name: string, v: unknown): number => {
    const n = Number(v)
    if (!Number.isFinite(n)) throw new Error(`${ctx}: ${name} is not a valid number`)
    return n
  }

  const setBox = (id: unknown, patch: unknown) => {
    const key = String(id)
    const el = guard('setBox', id)
    const p = (patch ?? {}) as Record<string, unknown>
    const cur = pending.get(key) ?? {
      id: key,
      x: el.x,
      y: el.y,
      w: el.w,
      h: el.h,
      rotation: el.rotation,
    }
    const num = (v: unknown, name: string, fallback: number): number => {
      if (v === undefined) return fallback
      const n = Number(v)
      if (!Number.isFinite(n)) throw new Error(`setBox("${key}"): ${name} is not a valid number`)
      return n
    }
    const next: LayoutOp = {
      id: key,
      x: num(p.x, 'x', cur.x),
      y: num(p.y, 'y', cur.y),
      w: num(p.w, 'w', cur.w),
      h: num(p.h, 'h', cur.h),
      rotation: num(p.rotation, 'rotation', cur.rotation),
      ...(el.groupId ? { groupId: el.groupId } : {}),
    }
    if (next.w < 1 || next.h < 1) throw new Error(`setBox("${key}"): w/h must be ≥ 1px`)
    pending.set(key, next)
  }

  /** Relative move: based on the script's view of the "current value" (stacked on pending changes if any) */
  const moveBy = (id: unknown, dx: unknown, dy: unknown) => {
    const key = String(id)
    const el = guard('moveBy', id)
    const base = pending.get(key) ?? el
    setBox(id, {
      x: base.x + reqNum(`moveBy("${key}")`, 'dx', dx),
      y: base.y + reqNum(`moveBy("${key}")`, 'dy', dy),
    })
  }

  /** Relative resize: like moveBy, result no smaller than 1px */
  const resizeBy = (id: unknown, dw: unknown, dh: unknown) => {
    const key = String(id)
    const el = guard('resizeBy', id)
    const base = pending.get(key) ?? el
    setBox(id, {
      w: Math.max(1, base.w + reqNum(`resizeBy("${key}")`, 'dw', dw)),
      h: Math.max(1, base.h + reqNum(`resizeBy("${key}")`, 'dh', dh)),
    })
  }

  /** String/paragraph array → EditParagraph[] (strings split into paragraphs by \n, one run per paragraph) */
  const normalizeParagraphs = (ctx: string, v: unknown): EditParagraph[] => {
    if (typeof v === 'string') {
      return v.split('\n').map((line) => ({ runs: [{ text: line }] }))
    }
    if (!Array.isArray(v) || v.length === 0) {
      throw new Error(`${ctx}: the argument must be a string or a non-empty paragraph array`)
    }
    return v.map((item, i) => {
      if (typeof item === 'string') return { runs: [{ text: item }] }
      const p = (item ?? {}) as Record<string, unknown>
      // Paragraphs already in {runs:[…]} shape pass through (minimal validation only)
      if (Array.isArray(p.runs)) return item as EditParagraph
      if (typeof p.text !== 'string')
        throw new Error(`${ctx}: paragraph ${i + 1} is missing a text string`)
      const fontSize =
        p.fontSize === undefined
          ? undefined
          : reqNum(ctx, `paragraph ${i + 1} fontSize`, p.fontSize)
      return {
        runs: [
          {
            text: p.text,
            ...(p.bold ? { bold: true } : {}),
            ...(p.italic ? { italic: true } : {}),
            ...(p.underline ? { underline: true } : {}),
            ...(fontSize !== undefined ? { fontSize } : {}),
            ...(typeof p.fontFamily === 'string' && p.fontFamily
              ? { fontFamily: p.fontFamily }
              : {}),
            ...(typeof p.color === 'string' && p.color
              ? { color: normalizeHex(p.color) ?? p.color }
              : {}),
          },
        ],
        ...(p.align === 'left' || p.align === 'center' || p.align === 'right'
          ? { align: p.align }
          : {}),
      }
    })
  }

  /** Replace an element's text wholesale: string or paragraph array (same format as apply_ops setText paragraphs) */
  const setText = (id: unknown, textOrParagraphs: unknown) => {
    const key = String(id)
    const el = guard('setText', id)
    edits.push({
      kind: 'text',
      id: key,
      paragraphs: normalizeParagraphs(`setText("${key}")`, textOrParagraphs),
      ...(el.groupId ? { groupId: el.groupId } : {}),
    })
  }

  /** Change style without changing text: pass only the fields to change, applied to all the element's text */
  const setStyle = (id: unknown, patch: unknown) => {
    const key = String(id)
    const el = guard('setStyle', id)
    const p = (patch ?? {}) as Record<string, unknown>
    const style: SlideStylePatch = {}
    if (p.fontSize !== undefined) {
      const n = reqNum(`setStyle("${key}")`, 'fontSize', p.fontSize)
      if (n <= 0) throw new Error(`setStyle("${key}"): fontSize must be > 0`)
      style.fontSize = n
    }
    if (p.color !== undefined) {
      const hex = normalizeHex(p.color)
      if (!hex) throw new Error(`setStyle("${key}"): color must be #RRGGBB`)
      style.color = hex
    }
    if (p.bold !== undefined) style.bold = Boolean(p.bold)
    if (p.italic !== undefined) style.italic = Boolean(p.italic)
    if (p.underline !== undefined) style.underline = Boolean(p.underline)
    if (p.fontFamily !== undefined) {
      if (typeof p.fontFamily !== 'string' || !p.fontFamily)
        throw new Error(`setStyle("${key}"): fontFamily must be a non-empty string`)
      style.fontFamily = p.fontFamily
    }
    if (p.align !== undefined) {
      if (p.align !== 'left' && p.align !== 'center' && p.align !== 'right') {
        throw new Error(`setStyle("${key}"): align must be one of left/center/right`)
      }
      style.align = p.align
    }
    if (Object.keys(style).length === 0) {
      throw new Error(
        `setStyle("${key}"): pass at least one style field (fontSize/color/bold/italic/underline/fontFamily/align)`,
      )
    }
    edits.push({ kind: 'style', id: key, style, ...(el.groupId ? { groupId: el.groupId } : {}) })
  }

  /** Solid fill: #RRGGBB or 'none' */
  const setFill = (id: unknown, fill: unknown) => {
    const key = String(id)
    const el = guard('setFill', id)
    const grp = el.groupId ? { groupId: el.groupId } : {}
    if (fill === 'none') {
      edits.push({ kind: 'fill', id: key, fill: 'none', ...grp })
      return
    }
    const hex = normalizeHex(fill)
    if (!hex) throw new Error(`setFill("${key}"): fill must be #RRGGBB or 'none'`)
    edits.push({ kind: 'fill', id: key, fill: hex, ...grp })
  }

  /** Stroke: {color?,widthPt?} or null (remove the stroke) */
  const setStroke = (id: unknown, stroke: unknown) => {
    const key = String(id)
    const el = guard('setStroke', id)
    const grp = el.groupId ? { groupId: el.groupId } : {}
    if (stroke === null || stroke === undefined) {
      edits.push({ kind: 'stroke', id: key, stroke: null, ...grp })
      return
    }
    const p = stroke as Record<string, unknown>
    const hex = p.color === undefined ? '#000000' : normalizeHex(p.color)
    if (!hex) throw new Error(`setStroke("${key}"): color must be #RRGGBB`)
    const widthPt =
      p.widthPt === undefined ? 1 : reqNum(`setStroke("${key}")`, 'widthPt', p.widthPt)
    if (widthPt <= 0) throw new Error(`setStroke("${key}"): widthPt must be > 0`)
    edits.push({ kind: 'stroke', id: key, stroke: { color: hex, widthPt }, ...grp })
  }

  const log = (...args: unknown[]) => {
    if (logs.length >= 50) return
    logs.push(args.map((a) => (typeof a === 'string' ? a : JSON.stringify(a))).join(' '))
  }

  let returned: unknown
  try {
    returned = interpretLayoutScript(code, {
      els: elements,
      canvas,
      setBox,
      moveBy,
      resizeBy,
      setText,
      setStyle,
      setFill,
      setStroke,
      log,
    })
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err)
    return { ops: [], edits: [], logs, error: msg }
  }

  let returnedStr: string | undefined
  if (returned !== undefined) {
    try {
      returnedStr = typeof returned === 'string' ? returned : JSON.stringify(returned)
    } catch {
      returnedStr = String(returned)
    }
  }
  return {
    ops: [...pending.values()],
    edits,
    logs,
    ...(returnedStr !== undefined ? { returned: returnedStr } : {}),
  }
}

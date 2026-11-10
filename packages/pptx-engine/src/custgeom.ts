/**
 * <a:custGeom> custom geometry parsing (read-only rendering).
 *
 * fast-xml-parser does not preserve document order of mixed child nodes
 * (moveTo/lnTo/… order would be lost), so we scan tags in order directly over
 * the raw XML bytes.
 * Output is a normalized SVG path: coordinates 0..1, relative to each <a:path>'s
 * declared w/h; arcTo is converted to cubic beziers; gdLst/avLst formulas are
 * evaluated first so pt/arcTo can reference them.
 */
import type { CustomGeometry } from './types'

// Tag matcher tolerating '>' inside attribute values (same pattern as scan.ts)
const TAG_RE = /<\/?(?:[^<>"']|"[^"]*"|'[^']*')*>/g
const NAME_RE = /^<\/?\s*([A-Za-z_][\w:.-]*)/
const ATTR_RE = /([\w:]+)\s*=\s*"([^"]*)"/g

function tagAttrs(tag: string): Record<string, string> {
  const out: Record<string, string> = {}
  ATTR_RE.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = ATTR_RE.exec(tag))) out[m[1]!] = m[2]!
  return out
}

const DEG = Math.PI / 180
/** OOXML angle (1/60000 of a degree) → radians */
const a2r = (v: number) => (v / 60000) * DEG

/**
 * gd formula evaluation (CT_GeomGuide fmla). Evaluated in order; may reference
 * earlier guides and built-in variables (w/h/ss/hc/vc/l/t/r/b/wdN/hdN/ssdN/cdN…,
 * angles in 1/60000 of a degree).
 */
export function evalGuides(
  gds: Array<{ name: string; fmla: string }>,
  w: number,
  h: number,
): Record<string, number> {
  const ss = Math.min(w, h)
  const env: Record<string, number> = {
    w,
    h,
    ss,
    ls: Math.max(w, h),
    hc: w / 2,
    vc: h / 2,
    l: 0,
    t: 0,
    r: w,
    b: h,
    wd2: w / 2,
    wd3: w / 3,
    wd4: w / 4,
    wd5: w / 5,
    wd6: w / 6,
    wd8: w / 8,
    wd10: w / 10,
    wd12: w / 12,
    wd32: w / 32,
    hd2: h / 2,
    hd3: h / 3,
    hd4: h / 4,
    hd5: h / 5,
    hd6: h / 6,
    hd8: h / 8,
    hd10: h / 10,
    hd12: h / 12,
    ssd2: ss / 2,
    ssd4: ss / 4,
    ssd6: ss / 6,
    ssd8: ss / 8,
    ssd16: ss / 16,
    ssd32: ss / 32,
    cd2: 10800000,
    cd4: 5400000,
    cd8: 2700000,
    '3cd4': 16200000,
    '3cd8': 8100000,
    '5cd8': 13500000,
    '7cd8': 18900000,
  }
  const val = (tok: string | undefined): number => {
    if (tok == null) return 0
    const n = Number(tok)
    return Number.isFinite(n) ? n : (env[tok] ?? 0)
  }
  for (const gd of gds) {
    const parts = gd.fmla.trim().split(/\s+/)
    const op = parts[0]
    const x = val(parts[1])
    const y = val(parts[2])
    const z = val(parts[3])
    let out: number
    switch (op) {
      case 'val':
        out = x
        break
      case '*/':
        out = z === 0 ? 0 : (x * y) / z
        break
      case '+-':
        out = x + y - z
        break
      case '+/':
        out = z === 0 ? 0 : (x + y) / z
        break
      case '?:':
        out = x > 0 ? y : z
        break
      case 'abs':
        out = Math.abs(x)
        break
      case 'min':
        out = Math.min(x, y)
        break
      case 'max':
        out = Math.max(x, y)
        break
      case 'pin':
        out = Math.min(Math.max(y, x), z)
        break
      case 'mod':
        out = Math.sqrt(x * x + y * y + z * z)
        break
      case 'sqrt':
        out = Math.sqrt(Math.max(x, 0))
        break
      case 'at2':
        out = (Math.atan2(y, x) / DEG) * 60000
        break
      case 'cat2':
        out = x * Math.cos(Math.atan2(z, y))
        break
      case 'sat2':
        out = x * Math.sin(Math.atan2(z, y))
        break
      case 'cos':
        out = x * Math.cos(a2r(y))
        break
      case 'sin':
        out = x * Math.sin(a2r(y))
        break
      case 'tan':
        out = x * Math.tan(a2r(y))
        break
      default:
        out = 0
    }
    env[gd.name] = out
  }
  return env
}

interface RawCmd {
  c: 'M' | 'L' | 'Q' | 'C' | 'Z' | 'A'
  /** pt coordinate tokens for M/L/Q/C (numbers or guide names) */
  pts?: Array<[string, string]>
  /** arcTo attribute tokens */
  arc?: { wR: string; hR: string; stAng: string; swAng: string }
}

interface RawPath {
  w?: number
  h?: number
  fill?: string
  stroke?: string
  cmds: RawCmd[]
}

const CMD_PT_COUNT: Record<string, number> = {
  'a:moveTo': 1,
  'a:lnTo': 1,
  'a:quadBezTo': 2,
  'a:cubicBezTo': 3,
}
const CMD_LETTER: Record<string, RawCmd['c']> = {
  'a:moveTo': 'M',
  'a:lnTo': 'L',
  'a:quadBezTo': 'Q',
  'a:cubicBezTo': 'C',
}

/**
 * Extract and parse <a:custGeom> from a shape's raw XML.
 * shapeW/shapeH: element box in EMU (basis for guide evaluation; fallback space
 * when a path declares no w/h).
 */
export function parseCustGeom(
  shapeXml: string,
  shapeW: number,
  shapeH: number,
): CustomGeometry | undefined {
  const start = shapeXml.indexOf('<a:custGeom')
  if (start < 0) return undefined
  const end = shapeXml.indexOf('</a:custGeom>', start)
  if (end < 0) return undefined
  const xml = shapeXml.slice(start, end)

  const gds: Array<{ name: string; fmla: string }> = []
  const paths: RawPath[] = []
  let inGuides = false
  let cur: RawPath | null = null
  let pending: { c: RawCmd['c']; need: number; pts: Array<[string, string]> } | null = null

  TAG_RE.lastIndex = 0
  let m: RegExpExecArray | null
  while ((m = TAG_RE.exec(xml))) {
    const tag = m[0]
    if (tag.startsWith('<!--') || tag.startsWith('<![') || tag.startsWith('<?')) continue
    const closing = tag.startsWith('</')
    const name = NAME_RE.exec(tag)?.[1] ?? ''
    if (closing) {
      if (name === 'a:avLst' || name === 'a:gdLst') inGuides = false
      else if (name === 'a:path' && cur) {
        paths.push(cur)
        cur = null
        pending = null
      }
      continue
    }
    const self = tag.endsWith('/>')
    switch (name) {
      case 'a:avLst':
      case 'a:gdLst':
        if (!self) inGuides = true
        break
      case 'a:gd': {
        if (!inGuides) break // gd elsewhere (e.g. cxnLst) does not participate
        const a = tagAttrs(tag)
        if (a.name && a.fmla) gds.push({ name: a.name, fmla: a.fmla })
        break
      }
      case 'a:path': {
        const a = tagAttrs(tag)
        cur = {
          w: a.w ? Number(a.w) || 0 : undefined,
          h: a.h ? Number(a.h) || 0 : undefined,
          fill: a.fill,
          stroke: a.stroke,
          cmds: [],
        }
        if (self) {
          paths.push(cur)
          cur = null
        }
        break
      }
      case 'a:moveTo':
      case 'a:lnTo':
      case 'a:quadBezTo':
      case 'a:cubicBezTo':
        if (cur) pending = { c: CMD_LETTER[name]!, need: CMD_PT_COUNT[name]!, pts: [] }
        break
      case 'a:pt': {
        if (!cur || !pending) break
        const a = tagAttrs(tag)
        pending.pts.push([a.x ?? '0', a.y ?? '0'])
        if (pending.pts.length >= pending.need) {
          cur.cmds.push({ c: pending.c, pts: pending.pts })
          pending = null
        }
        break
      }
      case 'a:arcTo': {
        if (!cur) break
        const a = tagAttrs(tag)
        cur.cmds.push({
          c: 'A',
          arc: { wR: a.wR ?? '0', hR: a.hR ?? '0', stAng: a.stAng ?? '0', swAng: a.swAng ?? '0' },
        })
        break
      }
      case 'a:close':
        if (cur) cur.cmds.push({ c: 'Z' })
        break
    }
  }

  const env = evalGuides(gds, shapeW, shapeH)
  const resolve = (tok: string): number => {
    const n = Number(tok)
    return Number.isFinite(n) ? n : (env[tok] ?? 0)
  }

  const absList = paths.map((p) => toAbsCmds(p, resolve))
  // Fallback normalization space when a path declares no w/h and the element box is 0:
  // shared across paths (they live in one guide coordinate space; per-path maxima would skew relative scale)
  let fw = 0
  let fh = 0
  paths.forEach((p, i) => {
    if (!p.w) fw = Math.max(fw, maxCoord(absList[i]!, 0))
    if (!p.h) fh = Math.max(fh, maxCoord(absList[i]!, 1))
  })

  // Each path normalized independently (its w/h maps to the full element box), emitted
  // in document order; the three output layers keep that order internally (cross-layer
  // paint order is bounded by the render contract's fixed path/fillPath/strokePath sequence)
  const buckets = { path: [] as string[], fillPath: [] as string[], strokePath: [] as string[] }
  paths.forEach((p, i) => {
    const abs = absList[i]!
    if (!abs.length) return
    const vw = p.w || shapeW || fw || 1
    const vh = p.h || shapeH || fh || 1
    const d = emitNorm(abs, vw, vh)
    if (!d) return
    const fillNone = p.fill === 'none'
    const strokeNone = p.stroke === '0' || p.stroke === 'false' || p.stroke === 'none'
    if (fillNone && strokeNone) return
    if (fillNone) buckets.strokePath.push(d)
    else if (strokeNone) buckets.fillPath.push(d)
    else buckets.path.push(d)
  })

  const out: CustomGeometry = {}
  if (buckets.path.length) out.path = buckets.path.join(' ')
  if (buckets.fillPath.length) out.fillPath = buckets.fillPath.join(' ')
  if (buckets.strokePath.length) out.strokePath = buckets.strokePath.join(' ')
  return out.path || out.fillPath || out.strokePath ? out : undefined
}

interface AbsCmd {
  c: 'M' | 'L' | 'Q' | 'C' | 'Z'
  xy: number[]
}

/**
 * Ray angle → ellipse parametric angle, staying in the same revolution.
 * OOXML arcTo angles point at the ellipse from the center; the parametric form
 * (wR·cos t, hR·sin t) only matches when wR == hR.
 */
function paramAngle(a: number, wR: number, hR: number): number {
  if (wR === hR || !wR || !hR) return a
  const t = Math.atan2(Math.sin(a) * wR, Math.cos(a) * hR)
  return t + 2 * Math.PI * Math.round((a - t) / (2 * Math.PI))
}

/** Commands → absolute coordinates (path space); arcTo converted to cubics in place. */
function toAbsCmds(p: RawPath, resolve: (tok: string) => number): AbsCmd[] {
  const out: AbsCmd[] = []
  let cx = 0
  let cy = 0
  let sx = 0
  let sy = 0
  for (const cmd of p.cmds) {
    if (cmd.c === 'Z') {
      out.push({ c: 'Z', xy: [] })
      // close returns the pen to the subpath start; a following arcTo derives its center from here
      cx = sx
      cy = sy
      continue
    }
    if (cmd.c === 'A') {
      const wR = resolve(cmd.arc!.wR)
      const hR = resolve(cmd.arc!.hR)
      const stRay = a2r(resolve(cmd.arc!.stAng))
      const swRay = a2r(resolve(cmd.arc!.swAng))
      if (swRay === 0) continue
      const st = paramAngle(stRay, wR, hR)
      // Endpoint-converted sweep, restored to swRay's revolution count (each endpoint shifts < π/2)
      let sw = paramAngle(stRay + swRay, wR, hR) - st
      sw += 2 * Math.PI * Math.round((swRay - sw) / (2 * Math.PI))
      // Repeated full revolutions retrace the same ellipse. Keep one loop and
      // the final partial sweep so a single arc cannot expand without bound.
      const fullTurn = 2 * Math.PI
      if (Math.abs(sw) > fullTurn) {
        sw = Math.sign(sw) * (fullTurn + (Math.abs(sw) % fullTurn))
      }
      // Treat current point as at stAng on the ellipse and derive the center
      const ecx = cx - wR * Math.cos(st)
      const ecy = cy - hR * Math.sin(st)
      const segs = Math.max(1, Math.ceil(Math.abs(sw) / (Math.PI / 2)))
      const da = sw / segs
      const k = (4 / 3) * Math.tan(da / 4)
      for (let i = 0; i < segs; i++) {
        const a1 = st + i * da
        const a2 = a1 + da
        const x1 = ecx + wR * Math.cos(a1)
        const y1 = ecy + hR * Math.sin(a1)
        const x2 = ecx + wR * Math.cos(a2)
        const y2 = ecy + hR * Math.sin(a2)
        out.push({
          c: 'C',
          xy: [
            x1 - k * wR * Math.sin(a1),
            y1 + k * hR * Math.cos(a1),
            x2 + k * wR * Math.sin(a2),
            y2 - k * hR * Math.cos(a2),
            x2,
            y2,
          ],
        })
        cx = x2
        cy = y2
      }
      continue
    }
    const xy: number[] = []
    for (const [xs, ys] of cmd.pts!) {
      xy.push(resolve(xs), resolve(ys))
    }
    out.push({ c: cmd.c, xy })
    if (cmd.c === 'M') {
      sx = xy[0]!
      sy = xy[1]!
    }
    cx = xy[xy.length - 2]!
    cy = xy[xy.length - 1]!
  }
  return out
}

function maxCoord(cmds: AbsCmd[], axis: 0 | 1): number {
  let mx = 0
  for (const c of cmds) for (let i = axis; i < c.xy.length; i += 2) mx = Math.max(mx, c.xy[i]!)
  return mx
}

function emitNorm(cmds: AbsCmd[], vw: number, vh: number): string {
  const r5 = (v: number) => Math.round(v * 100000) / 100000
  const parts: string[] = []
  for (const c of cmds) {
    parts.push(c.c)
    for (let i = 0; i < c.xy.length; i += 2) {
      parts.push(String(r5(c.xy[i]! / vw)), String(r5(c.xy[i + 1]! / vh)))
    }
  }
  return parts.length ? parts.join(' ') : ''
}

/** Absolute path command in <a:path> space (EMU when w/h are the element extents). */
export interface CustGeomPathCmd {
  op: 'M' | 'L' | 'C' | 'Q' | 'Z'
  pts: number[]
}

/** Editable single-path custom geometry: what Edit Points writes back as <a:custGeom>. */
export interface CustGeomPath {
  w: number
  h: number
  cmds: CustGeomPathCmd[]
}

export const CUST_GEOM_PT_COUNT: Record<CustGeomPathCmd['op'], number> = {
  M: 2,
  L: 2,
  C: 6,
  Q: 4,
  Z: 0,
}

export function validCustGeomPath(p: unknown): p is CustGeomPath {
  if (!p || typeof p !== 'object') return false
  const { w, h, cmds } = p as Partial<CustGeomPath>
  if (!Number.isFinite(w) || !Number.isFinite(h) || w! <= 0 || h! <= 0) return false
  if (!Array.isArray(cmds) || !cmds.length || cmds[0]!.op !== 'M') return false
  return cmds.every(
    (c) =>
      c &&
      typeof c === 'object' &&
      c.op in CUST_GEOM_PT_COUNT &&
      Array.isArray(c.pts) &&
      c.pts.length === CUST_GEOM_PT_COUNT[c.op] &&
      c.pts.every((v) => Number.isFinite(v)),
  )
}

const CMD_TAG: Record<Exclude<CustGeomPathCmd['op'], 'Z'>, string> = {
  M: 'a:moveTo',
  L: 'a:lnTo',
  C: 'a:cubicBezTo',
  Q: 'a:quadBezTo',
}

/** Serialize a path as a complete <a:custGeom> (integer coordinates, path w/h as the space). */
export function custGeomXml(p: CustGeomPath): string {
  const body = p.cmds
    .map((c) => {
      if (c.op === 'Z') return '<a:close/>'
      let pts = ''
      for (let i = 0; i < c.pts.length; i += 2)
        pts += `<a:pt x="${Math.round(c.pts[i]!)}" y="${Math.round(c.pts[i + 1]!)}"/>`
      return `<${CMD_TAG[c.op]}>${pts}</${CMD_TAG[c.op]}>`
    })
    .join('')
  return (
    `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>` +
    `<a:pathLst><a:path w="${Math.round(p.w)}" h="${Math.round(p.h)}">${body}</a:path></a:pathLst></a:custGeom>`
  )
}

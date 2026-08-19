/**
 * Hierarchy-family (orgChart1) diagram layout interpreter.
 *
 * PowerPoint org charts carry the full layout algorithm in diagrams/layoutN.xml:
 * box proportions and gaps as <dgm:constr> factors, and the std/hang branch
 * decision as <dgm:choose> conditions on node depth. This module evaluates those
 * instead of approximating them:
 *  - box h = `h` fact × w (0.5); sibling gap = `sibSp` fact × w (0.21);
 *    row gap = `sp` fact × w (0.21) — all read from the constraint list
 *  - subtrees pack by per-row contours (a hanging overhang may slide under a
 *    sibling's box row), parent centered on the midpoint of its outer children
 *  - `init` auto-branch: deep all-leaf branches hang beside the trunk once the
 *    standard layout reaches 4+ rows (see branchOf)
 *  - hang trunk drops from the hidden rootConnector (left 20% of the box, so
 *    x = left + 0.1w); hanging boxes start at left + 0.25w (PowerPoint measured)
 *  - the whole unit-space layout scales uniformly into the frame and centers
 * All geometry below is in units of one box width.
 */

export interface HierTreeNode {
  texts: string[]
  children: HierTreeNode[]
  asst?: boolean
  hierBranch?: string
  spPr?: any
}

export interface HierConstraints {
  boxAspect: number
  sibSp: number
  sp: number
  trunkOff: number
  hangIndent: number
  fontMax: number
}

export interface HierBox {
  node: HierTreeNode
  x: number
  y: number
  w: number
  h: number
}

export interface HierLine {
  x1: number
  y1: number
  x2: number
  y2: number
}

const DEFAULTS: HierConstraints = {
  boxAspect: 0.5,
  sibSp: 0.21,
  sp: 0.21,
  trunkOff: 0.1,
  hangIndent: 0.25,
  fontMax: 65,
}

/** Pull the factors this engine consumes out of the layout part's constraint lists. */
export function parseHierConstraints(layoutXml: string | undefined): HierConstraints {
  const cons = { ...DEFAULTS }
  if (!layoutXml) return cons
  const fact = (re: RegExp): number | undefined => {
    const m = re.exec(layoutXml)
    if (!m) return undefined
    const v = parseFloat(m[1]!)
    return Number.isFinite(v) && v > 0 ? v : undefined
  }
  cons.boxAspect =
    fact(
      /<dgm:constr type="h" for="des" forName="rootComposite1?"[^>]*refForName="rootComposite1?"[^>]*fact="([\d.]+)"/,
    ) ?? cons.boxAspect
  cons.sibSp =
    fact(/<dgm:constr type="sibSp"[^>]*refForName="rootComposite1?"[^>]*fact="([\d.]+)"/) ??
    cons.sibSp
  cons.sp =
    fact(/<dgm:constr type="sp" for="des" forName="hierRoot1"[^>]*fact="([\d.]+)"/) ?? cons.sp
  const connW = fact(/<dgm:constr type="w" for="ch" forName="rootConnector1?"[^>]*fact="([\d.]+)"/)
  if (connW != null) cons.trunkOff = connW / 2
  const fsz = /<dgm:constr type="primFontSz"[^>]*\bval="([\d.]+)"/.exec(layoutXml)
  if (fsz) cons.fontMax = parseFloat(fsz[1]!) || cons.fontMax
  return cons
}

interface Sub {
  rows: number
  /** per-row x extents relative to this subtree's root box center */
  l: number[]
  r: number[]
  boxes: Array<{ node: HierTreeNode; cx: number; row: number }>
  lines: HierLine[]
}

const isLeafAsst = (c: HierTreeNode) => !!c.asst && !c.children.length
const kidsOf = (n: HierTreeNode) => n.children.filter((c) => !isLeafAsst(c))
const asstsOf = (n: HierTreeNode) => n.children.filter(isLeafAsst)

type Branch = 'std' | 'hangR' | 'hangL'

/** Whole-tree row count with every branch laid out standard (incl. assistant rows). */
function stdRowsOf(node: HierTreeNode): number {
  const kids = kidsOf(node)
  const asstRows = asstsOf(node).length ? 1 : 0
  return kids.length ? 1 + asstRows + Math.max(...kids.map(stdRowsOf)) : 1 + asstRows
}

/**
 * Branch style for a node's children. Explicit hierBranch (from the node's
 * presentation point) wins. `init` means the algorithm decides; PowerPoint
 * measured (smartart-org-chart2 hangs its depth-3 leaves, smartart-recursion —
 * same depth, one row shorter — keeps everything standard): only charts whose
 * standard layout reaches 4+ rows hang their deep all-leaf branches, never at
 * the root row.
 */
function branchOf(
  node: HierTreeNode,
  depth: number,
  kids: HierTreeNode[],
  stdRows: number,
): Branch {
  if (!kids.length) return 'std'
  const allLeaves = node.children.every((c) => !c.children.length)
  const hb = node.hierBranch
  if (hb === 'l') return 'hangL'
  if (hb === 'r' || hb === 'hang') return 'hangR'
  if (hb === 'std') return 'std'
  return stdRows >= 4 && depth >= 2 && kids.length >= 2 && allLeaves ? 'hangR' : 'std'
}

function shiftSub(sub: Sub, dx: number, dRow: number, into: Sub, dy: number) {
  for (let i = 0; i < sub.rows; i++) {
    const row = dRow + i
    const sl = sub.l[i]! + dx
    const sr = sub.r[i]! + dx
    into.l[row] = into.l[row] == null ? sl : Math.min(into.l[row]!, sl)
    into.r[row] = into.r[row] == null ? sr : Math.max(into.r[row]!, sr)
  }
  for (const b of sub.boxes) into.boxes.push({ node: b.node, cx: b.cx + dx, row: b.row + dRow })
  for (const ln of sub.lines)
    into.lines.push({ x1: ln.x1 + dx, y1: ln.y1 + dy, x2: ln.x2 + dx, y2: ln.y2 + dy })
  into.rows = Math.max(into.rows, dRow + sub.rows)
}

export function layoutHierTree(
  roots: HierTreeNode[],
  cons: HierConstraints,
  frameCx: number,
  frameCy: number,
): {
  boxes: HierBox[]
  lines: Array<{ x: number; y: number; cx: number; cy: number }>
  boxW: number
} | null {
  if (!roots.length) return null
  const bh = cons.boxAspect
  const pitch = bh + cons.sp
  const rowTop = (row: number) => row * pitch
  const stdRows = Math.max(...roots.map(stdRowsOf))

  const layout = (node: HierTreeNode, depth: number): Sub => {
    const kids = kidsOf(node)
    const assts = asstsOf(node)
    const out: Sub = { rows: 1, l: [-0.5], r: [0.5], boxes: [{ node, cx: 0, row: 0 }], lines: [] }
    if (!kids.length && !assts.length) return out

    const branch = branchOf(node, depth, kids, stdRows)
    const mir = branch === 'hangL' ? -1 : 1
    const trunkX = branch === 'std' ? 0 : mir * (cons.trunkOff - 0.5)
    const childRow0 = 1 + (assts.length ? 1 : 0)

    if (assts.length) {
      // Leaf assistants: own row between the parent and its children, tucked
      // beside the trunk with a sp/2 elbow (PowerPoint measured)
      const asstMidY = rowTop(1) + bh / 2
      assts.forEach((a, j) => {
        const right = trunkX - cons.sp / 2 - j * (1 + cons.sibSp)
        out.boxes.push({ node: a, cx: right - 0.5, row: 1 })
        out.lines.push({ x1: right, y1: asstMidY, x2: trunkX, y2: asstMidY })
        out.l[1] = Math.min(out.l[1] ?? Infinity, right - 1)
        out.r[1] = Math.max(out.r[1] ?? -Infinity, trunkX)
      })
      out.rows = 2
    }

    if (!kids.length) {
      // assistant-only: trunk still drops to the assistant elbow
      out.lines.push({ x1: trunkX, y1: bh, x2: trunkX, y2: rowTop(1) + bh / 2 })
      return out
    }

    const subs = kids.map((k) => layout(k, depth + 1))

    if (branch === 'std') {
      // Contour packing: each subtree shifts right until every shared row keeps
      // a sibSp gap; the parent centers on its first/last child box centers
      const merged: { l: number[]; r: number[] } = { l: [], r: [] }
      const offs: number[] = []
      for (const sub of subs) {
        let dx = 0
        if (offs.length) {
          dx = -Infinity
          for (let i = 0; i < sub.rows; i++)
            if (merged.r[i] != null) dx = Math.max(dx, merged.r[i]! + cons.sibSp - sub.l[i]!)
          if (!Number.isFinite(dx)) dx = merged.r[0]! + cons.sibSp - sub.l[0]!
        }
        offs.push(dx)
        for (let i = 0; i < sub.rows; i++) {
          merged.l[i] =
            merged.l[i] == null ? sub.l[i]! + dx : Math.min(merged.l[i]!, sub.l[i]! + dx)
          merged.r[i] =
            merged.r[i] == null ? sub.r[i]! + dx : Math.max(merged.r[i]!, sub.r[i]! + dx)
        }
      }
      const mid = (offs[0]! + offs[offs.length - 1]!) / 2
      const busY = rowTop(childRow0) - cons.sp / 2
      out.lines.push({ x1: trunkX, y1: bh, x2: trunkX, y2: busY })
      for (let i = 0; i < subs.length; i++) {
        const cx = offs[i]! - mid
        out.lines.push({ x1: trunkX, y1: busY, x2: cx, y2: busY })
        out.lines.push({ x1: cx, y1: busY, x2: cx, y2: rowTop(childRow0) })
        shiftSub(subs[i]!, cx, childRow0, out, rowTop(childRow0))
      }
    } else {
      // Hanging: children stack below the trunk, boxes starting at parent-left
      // + hangIndent (mirrored for hangL); each child may itself be a subtree
      let row = childRow0
      let lastMidY = bh
      for (const sub of subs) {
        const cx = mir * (cons.hangIndent - 0.5 + 0.5)
        const midY = rowTop(row) + bh / 2
        out.lines.push({ x1: trunkX, y1: midY, x2: cx - mir * 0.5, y2: midY })
        shiftSub(sub, cx, row, out, rowTop(row))
        lastMidY = midY
        row += sub.rows
      }
      out.lines.push({ x1: trunkX, y1: bh, x2: trunkX, y2: lastMidY })
    }
    return out
  }

  // Top level (hierChild1): root subtrees pack like siblings
  const top: Sub = { rows: 0, l: [], r: [], boxes: [], lines: [] }
  const merged: { l: number[]; r: number[] } = { l: [], r: [] }
  let placed = 0
  for (const r of roots) {
    const sub = layout(r, 1)
    let dx = 0
    if (placed) {
      dx = -Infinity
      for (let i = 0; i < sub.rows; i++)
        if (merged.r[i] != null) dx = Math.max(dx, merged.r[i]! + cons.sibSp - sub.l[i]!)
      if (!Number.isFinite(dx)) dx = 0
    }
    for (let i = 0; i < sub.rows; i++) {
      merged.l[i] = merged.l[i] == null ? sub.l[i]! + dx : Math.min(merged.l[i]!, sub.l[i]! + dx)
      merged.r[i] = merged.r[i] == null ? sub.r[i]! + dx : Math.max(merged.r[i]!, sub.r[i]! + dx)
    }
    shiftSub(sub, dx, 0, top, 0)
    placed++
  }

  let minX = Infinity
  let maxX = -Infinity
  for (let i = 0; i < top.rows; i++) {
    if (top.l[i] != null) minX = Math.min(minX, top.l[i]!)
    if (top.r[i] != null) maxX = Math.max(maxX, top.r[i]!)
  }
  for (const ln of top.lines) {
    minX = Math.min(minX, ln.x1, ln.x2)
    maxX = Math.max(maxX, ln.x1, ln.x2)
  }
  const unitW = maxX - minX
  const unitH = top.rows * bh + (top.rows - 1) * cons.sp
  if (!(unitW > 0) || !(unitH > 0)) return null
  const scale = Math.min(frameCx / unitW, frameCy / unitH)
  const offX = (frameCx - unitW * scale) / 2 - minX * scale
  const offY = (frameCy - unitH * scale) / 2

  const boxes: HierBox[] = top.boxes.map((b) => ({
    node: b.node,
    x: offX + (b.cx - 0.5) * scale,
    y: offY + rowTop(b.row) * scale,
    w: scale,
    h: bh * scale,
  }))
  const lines = top.lines.map((ln) => ({
    x: offX + Math.min(ln.x1, ln.x2) * scale,
    y: offY + Math.min(ln.y1, ln.y2) * scale,
    cx: Math.abs(ln.x2 - ln.x1) * scale,
    cy: Math.abs(ln.y2 - ln.y1) * scale,
  }))
  return { boxes, lines, boxW: scale }
}

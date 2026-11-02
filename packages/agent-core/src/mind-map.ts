import type { AgentToolDef, MindMapTreeNode } from './types'

/**
 * Mind map generation shared by the docs and markdown assistants: the tool
 * definition the model calls with the whole tree in one shot, strict-but-
 * forgiving parsing of that tree, and a pure tidy-tree layout the card
 * component renders. No runtime dependencies — layout is plain arithmetic,
 * rendering is SVG.
 *
 * Two orientations are available. `top-down` puts the root on top and lets
 * the branches fan out below it; `left-right` puts the root in the middle
 * and splits the branches into a left and a right half, the way a
 * hand-drawn mind map reads. Both produce the same node boxes, so the card
 * and the exporters stay orientation-agnostic.
 */

/** Hard limits so a runaway model cannot blow up the card (or the tool call). */
export const MIND_MAP_LIMITS = {
  maxNodes: 120,
  maxDepth: 5,
  maxLabelChars: 60,
  maxFocusChars: 80,
} as const

/** Layout metrics (px). Node boxes are fixed-size; labels wrap to 2 lines. */
export const MIND_MAP_METRICS = {
  nodeWidth: 170,
  nodeHeight: 34,
  depthGap: 40,
  siblingGap: 14,
  padding: 10,
} as const

/** Fixed palette for PNG export: exports are artifacts, identical in both themes. */
export const MIND_MAP_EXPORT_COLORS = {
  background: '#ffffff',
  nodeFill: '#ffffff',
  nodeStroke: '#d0d3d8',
  rootFill: '#e8663f',
  rootStroke: '#d4552f',
  nodeText: '#242424',
  rootText: '#ffffff',
  edge: '#c8ccd2',
  badgeFill: '#0f7fff',
  badgeText: '#ffffff',
  toggleFill: '#f0f2f5',
  toggleStroke: '#d0d3d8',
  toggleText: '#606366',
} as const

/** Bounded recursion: the schema inlines down to the enforced depth limit; the
 *  deepest level carries a leaf-node shape, because Gemini's function schema
 *  refuses an array without `items` for the whole request (genoffice#2059). */
const nodeSchema = (depth: number): Record<string, unknown> => ({
  type: 'object',
  properties: {
    label: {
      type: 'string',
      description:
        "short node label (2-8 words) in the document's language, taken from its headings or key phrases",
    },
    block: {
      type: 'integer',
      description:
        '0-based block index this node summarizes (from the current block list); omit when the node spans several blocks',
    },
    children:
      depth > 0
        ? {
            type: 'array',
            description: 'sub-topics; 2-5 per node reads best',
            items: nodeSchema(depth - 1),
          }
        : {
            type: 'array',
            description: 'sub-topics; 2-5 per node reads best',
            items: {
              type: 'object',
              properties: {
                label: {
                  type: 'string',
                  description: "short node label (2-8 words) in the document's language",
                },
              },
              required: ['label'],
            },
          },
  },
  required: ['label'],
})

/**
 * The model emits the whole tree as the tool arguments — one request, no second
 * model round-trip. `focus` mirrors NotebookLM's customization prompt: the same
 * document can produce different maps for different angles.
 */
export const MIND_MAP_TOOL: AgentToolDef = {
  name: 'generate_mindmap',
  description:
    'Render an interactive mind map of the document (or of one focused aspect) as a diagram card in the chat. ' +
    'Call it when the user asks for a mind map / knowledge structure / overview diagram of the document or a topic. ' +
    'Pass the entire tree in this single call: the root is the document title or the focus topic, its children are the main sections, and so on. ' +
    "Labels must be short phrases in the document's own language and wording — never invent topics the document does not cover. " +
    'Cite the block index each node summarizes in `block` so the user can jump to the passage. ' +
    'This tool does not modify the document; after it returns, reply with one short sentence.',
  inputSchema: {
    type: 'object',
    properties: {
      focus: {
        type: 'string',
        description:
          'optional: the aspect the map should center on (e.g. "risk controls", "for beginners"); omit for a whole-document map',
      },
      root: {
        type: 'object',
        description: 'the mind map tree: root node with nested children',
        ...nodeSchema(MIND_MAP_LIMITS.maxDepth),
      },
    },
    required: ['root'],
  },
}

export type MindMapParseResult =
  | { ok: true; tree: MindMapTreeNode; nodeCount: number; depth: number; trimmed: boolean }
  | { ok: false; error: string }

/**
 * Validate the model's tree against the document: labels must be non-empty
 * strings, block citations must address real blocks (dropped, not fatal — the
 * model may cite from a stale context), and the size limits are enforced by
 * trimming. A tree without branches is rejected so the model retries with a
 * real structure instead of rendering a lonely root.
 */
export function parseMindMapTree(
  input: Record<string, unknown>,
  blockCount: number,
): MindMapParseResult {
  let nodeCount = 0
  let maxDepth = 0
  let trimmed = false

  const walk = (raw: unknown, depth: number): MindMapTreeNode | null => {
    if (nodeCount >= MIND_MAP_LIMITS.maxNodes || depth > MIND_MAP_LIMITS.maxDepth) {
      trimmed = true
      return null
    }
    if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) return null
    const obj = raw as Record<string, unknown>
    const label = typeof obj.label === 'string' ? obj.label.trim() : ''
    if (!label) return null
    nodeCount += 1
    maxDepth = Math.max(maxDepth, depth)
    const node: MindMapTreeNode = { label: label.slice(0, MIND_MAP_LIMITS.maxLabelChars) }
    const block = obj.block
    if (Number.isInteger(block) && (block as number) >= 0 && (block as number) < blockCount) {
      node.block = block as number
    }
    const kids = Array.isArray(obj.children) ? obj.children : []
    const children: MindMapTreeNode[] = []
    for (const kid of kids) {
      const placed = walk(kid, depth + 1)
      if (placed) children.push(placed)
    }
    if (children.length > 0) node.children = children
    return node
  }

  const root = walk(input.root, 0)
  if (!root) {
    return {
      ok: false,
      error:
        'root must be an object with a non-empty "label" string — pass the tree as {"root": {"label": "...", "children": [...]}}',
    }
  }
  if (!root.children || root.children.length === 0) {
    return {
      ok: false,
      error:
        'the map has no branches — expand the root into its main sections (2-5 children, each with its own children) and call again',
    }
  }
  return { ok: true, tree: root, nodeCount, depth: maxDepth, trimmed }
}

/** Total nodes in a tree (all levels). */
export function countMindMapNodes(root: MindMapTreeNode): number {
  let n = 1
  for (const child of root.children ?? []) n += countMindMapNodes(child)
  return n
}

/**
 * Serialize the tree as a mermaid `mindmap` diagram: one indented line per
 * node, root first. Mermaid reads `()[]{}` as node shapes, so they are
 * stripped from labels — a bare line is always a valid text node.
 */
export function mindMapToMermaid(root: MindMapTreeNode): string {
  const clean = (label: string): string => {
    const text = label
      .replace(/[[\]{}()"`]|%%|:::/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
    return text || 'Topic'
  }
  const lines: string[] = ['mindmap']
  const walk = (node: MindMapTreeNode, depth: number): void => {
    lines.push(`${'  '.repeat(depth + 1)}${clean(node.label)}`)
    for (const child of node.children ?? []) walk(child, depth + 1)
  }
  walk(root, 0)
  return lines.join('\n')
}

// ---- label wrapping (pure; CJK counts double width) ----

const CJK_RE = /[\u2E80-\u9FFF\uAC00-\uD7AF\uF900-\uFAFF\uFF00-\uFFEF\u3040-\u30FF]/

/** Display width units: CJK/full-width chars count 2, ASCII counts 1. */
export function labelUnits(text: string): number {
  let units = 0
  for (const ch of text) units += CJK_RE.test(ch) ? 2 : 1
  return units
}

interface Token {
  text: string
  units: number
}

/** Split a label into wrappable tokens: latin words keep trailing spaces, CJK splits per char. */
function tokenize(label: string): Token[] {
  const tokens: Token[] = []
  let word = ''
  const flush = () => {
    if (word) {
      tokens.push({ text: word, units: labelUnits(word) })
      word = ''
    }
  }
  for (const ch of label) {
    if (CJK_RE.test(ch)) {
      flush()
      tokens.push({ text: ch, units: 2 })
    } else if (ch === ' ') {
      word += ch
      flush()
    } else {
      word += ch
    }
  }
  flush()
  return tokens
}

/** Hard-split one oversized token into unit-bounded pieces. */
function splitToken(token: Token, maxUnits: number): Token[] {
  const pieces: Token[] = []
  let text = ''
  let units = 0
  for (const ch of token.text) {
    const chUnits = CJK_RE.test(ch) ? 2 : 1
    if (units + chUnits > maxUnits && text) {
      pieces.push({ text, units })
      text = ''
      units = 0
    }
    text += ch
    units += chUnits
  }
  if (text) pieces.push({ text, units })
  return pieces
}

/**
 * Greedy word/char wrap to at most `maxLines` lines of `maxUnits` display
 * units; the last line is ellipsized when content remains. Pure and total.
 */
export function wrapLabel(label: string, maxUnits: number, maxLines = 2): string[] {
  const safeMax = Math.max(4, maxUnits)
  const tokens: Token[] = []
  for (const token of tokenize(label)) {
    if (token.units > safeMax) tokens.push(...splitToken(token, safeMax))
    else tokens.push(token)
  }
  const lines: string[] = []
  let current = ''
  let currentUnits = 0
  let overflow = false
  for (const token of tokens) {
    if (currentUnits + token.units <= safeMax) {
      current += token.text
      currentUnits += token.units
      continue
    }
    if (lines.length + 1 >= maxLines) {
      overflow = true
      break
    }
    lines.push(current.trimEnd())
    current = token.text.trimStart()
    currentUnits = labelUnits(current)
  }
  if (current.trim()) lines.push(current.trimEnd())
  if (overflow && lines.length > 0) {
    const last = lines[lines.length - 1]
    lines[lines.length - 1] = fitEllipsis(last, safeMax)
  }
  return lines.length > 0 ? lines : [label]
}

/** Trim a line to the unit budget, leaving room for the ellipsis. */
function fitEllipsis(line: string, maxUnits: number): string {
  const ellipsisUnits = 2
  let out = ''
  let units = 0
  for (const ch of line) {
    const chUnits = CJK_RE.test(ch) ? 2 : 1
    if (units + chUnits + ellipsisUnits > maxUnits) break
    out += ch
    units += chUnits
  }
  return `${out}…`
}

// ---- tidy-tree layout ----

export interface MindMapLayoutNode {
  /** stable path id (child index path from the root, e.g. "0.2.1"); the collapse key */
  id: string
  label: string
  /** wrapped display lines (1-2) */
  lines: string[]
  /** cited document block index, when the model supplied a valid one */
  block?: number
  depth: number
  /** top-left of the node box */
  x: number
  y: number
  hasChildren: boolean
  /** direct children in the full tree (not the visible subset) */
  childCount: number
  /** nodes hidden below this one (subtree size minus itself) */
  descendants: number
  collapsed: boolean
}

export interface MindMapLayoutEdge {
  id: string
  /** SVG path `d`: cubic from the parent bottom border to the child top border */
  d: string
}

export interface MindMapLayout {
  nodes: MindMapLayoutNode[]
  edges: MindMapLayoutEdge[]
  width: number
  height: number
}

/**
 * How the tree is arranged on the canvas.
 *
 * - `top-down` — the root on top, the branches fanning out below it.
 * - `left-right` — the root in the middle, the branches split into a left and
 *   a right half, the way a hand-drawn mind map reads.
 */
export type MindMapOrientation = 'top-down' | 'left-right'

export interface MindMapLayoutOptions {
  /** defaults to `top-down` */
  orientation?: MindMapOrientation
}

/**
 * Vertical connector for the top-down layout: it leaves the parent at the
 * centre of its bottom border and arrives at the centre of the child's top
 * border, with the handles inside the gap so the curve never enters a box.
 */
function edgePath(parent: { x: number; y: number }, child: { x: number; y: number }): string {
  const px = parent.x + MIND_MAP_METRICS.nodeWidth / 2
  const py = parent.y + MIND_MAP_METRICS.nodeHeight
  const cx = child.x + MIND_MAP_METRICS.nodeWidth / 2
  const cy = child.y
  const k = Math.max(0, (cy - py) / 2)
  return `M ${px} ${py} C ${px} ${py + k}, ${cx} ${cy - k}, ${cx} ${cy}`
}

/**
 * Horizontal connector for the left-right layout: it leaves the parent at the
 * centre of the border that faces the child and arrives at the centre of the
 * child's facing border, with the handles inside the gap so the curve never
 * enters a box.
 *
 * `dir` is 1 when the parent sits to the left of the child (the right-hand
 * half of the map) and -1 when the parent sits to its right (the left half).
 */
function horizontalEdgePath(
  parent: { x: number; y: number },
  child: { x: number; y: number },
  dir: 1 | -1,
): string {
  const py = parent.y + MIND_MAP_METRICS.nodeHeight / 2
  const cy = child.y + MIND_MAP_METRICS.nodeHeight / 2
  const px = dir === 1 ? parent.x + MIND_MAP_METRICS.nodeWidth : parent.x
  const cx = dir === 1 ? child.x : child.x + MIND_MAP_METRICS.nodeWidth
  const k = Math.max(0, Math.abs(cx - px) / 2)
  return dir === 1
    ? `M ${px} ${py} C ${px + k} ${py}, ${cx - k} ${cy}, ${cx} ${cy}`
    : `M ${px} ${py} C ${px - k} ${py}, ${cx + k} ${cy}, ${cx} ${cy}`
}

/**
 * Lay the tree out. `top-down` maps depth to y and lets the leaves take
 * sequential x slots, with the parents centred over their children.
 * `left-right` puts the root in the middle, splits its branches into a left
 * and a right half, and maps depth to x inside each half while the leaves
 * take sequential y slots.
 *
 * Collapsed subtrees are skipped entirely (the node keeps a +N badge). Pure —
 * the same inputs always lay out the same.
 */
export function layoutMindMap(
  root: MindMapTreeNode,
  collapsed: ReadonlySet<string> = new Set(),
  options: MindMapLayoutOptions = {},
): MindMapLayout {
  if (options.orientation === 'left-right') return layoutMindMapLeftRight(root, collapsed)
  return layoutMindMapTopDown(root, collapsed)
}

/**
 * Top-down tidy tree: depth maps to y and leaves take sequential x slots, with
 * parents centered over their children.
 */
function layoutMindMapTopDown(
  root: MindMapTreeNode,
  collapsed: ReadonlySet<string>,
): MindMapLayout {
  const nodes: MindMapLayoutNode[] = []
  const edges: MindMapLayoutEdge[] = []
  const { padding } = MIND_MAP_METRICS
  let nextX = padding
  let maxDepth = 0

  const place = (node: MindMapTreeNode, id: string, depth: number): MindMapLayoutNode => {
    maxDepth = Math.max(maxDepth, depth)
    const kids = node.children ?? []
    const isCollapsed = collapsed.has(id)
    const placedKids = isCollapsed ? [] : kids.map((kid, i) => place(kid, `${id}.${i}`, depth + 1))
    let x: number
    if (placedKids.length > 0) {
      const first = placedKids[0]
      const last = placedKids[placedKids.length - 1]
      x = (first.x + last.x + MIND_MAP_METRICS.nodeWidth) / 2 - MIND_MAP_METRICS.nodeWidth / 2
    } else {
      x = nextX
      nextX += MIND_MAP_METRICS.nodeWidth + MIND_MAP_METRICS.siblingGap
    }
    const layoutNode: MindMapLayoutNode = {
      id,
      label: node.label,
      lines: wrapLabel(node.label, labelBudget(node.block !== undefined, kids.length > 0)),
      block: node.block,
      depth,
      x,
      y: depth * (MIND_MAP_METRICS.nodeHeight + MIND_MAP_METRICS.depthGap) + padding,
      hasChildren: kids.length > 0,
      childCount: kids.length,
      descendants: countMindMapNodes(node) - 1,
      collapsed: isCollapsed,
    }
    nodes.push(layoutNode)
    for (const kid of placedKids) {
      edges.push({ id: `${id}->${kid.id}`, d: edgePath(layoutNode, kid) })
    }
    return layoutNode
  }

  place(root, '0', 0)
  const { nodeWidth: w, nodeHeight: h, depthGap, siblingGap } = MIND_MAP_METRICS
  return {
    nodes,
    edges,
    width: Math.max(w, nextX - siblingGap - padding) + padding * 2,
    height: (maxDepth + 1) * (h + depthGap) - depthGap + padding * 2,
  }
}

/** Which half of the map a branch belongs to. */
type MindMapHalf = 'left' | 'right'

/**
 * Left-right tidy tree: the root sits in the middle, its branches are split
 * into a left and a right half, and inside each half depth maps to x while
 * the leaves take sequential y slots.
 *
 * The halves are balanced by leaf count, so a root with one heavy branch and
 * one light one does not push the whole map to one side. Each half is laid
 * out from the top on its own y slots, and the root is centred between the
 * outermost leaves.
 */
function layoutMindMapLeftRight(
  root: MindMapTreeNode,
  collapsed: ReadonlySet<string>,
): MindMapLayout {
  const { nodeWidth: W, nodeHeight: H, depthGap, siblingGap, padding } = MIND_MAP_METRICS
  const nodes: MindMapLayoutNode[] = []
  const edges: MindMapLayoutEdge[] = []

  const rootKids = root.children ?? []
  const rootCollapsed = collapsed.has('0')

  // ---- 1. split the root's branches into a left and a right half ----
  /** Leaves a branch contributes to the balance; a collapsed one counts as one. */
  const leafWeight = (node: MindMapTreeNode, id: string): number => {
    if (collapsed.has(id)) return 1
    const kids = node.children ?? []
    if (kids.length === 0) return 1
    return kids.reduce((sum, kid, i) => sum + leafWeight(kid, `${id}.${i}`), 0)
  }

  interface Branch {
    node: MindMapTreeNode
    id: string
    half: MindMapHalf
  }
  const left: Branch[] = []
  const right: Branch[] = []
  let leftWeight = 0
  let rightWeight = 0
  if (!rootCollapsed) {
    rootKids.forEach((kid, i) => {
      const id = `0.${i}`
      const weight = leafWeight(kid, id)
      // the first branch goes right, like the hand-drawn maps people expect
      if (leftWeight < rightWeight) {
        left.push({ node: kid, id, half: 'left' })
        leftWeight += weight
      } else {
        right.push({ node: kid, id, half: 'right' })
        rightWeight += weight
      }
    })
  }

  // ---- 2. lay each half out on its own y slots ----
  interface Placed {
    node: MindMapTreeNode
    id: string
    depth: number
    half: MindMapHalf
    y: number
  }
  const placed: Placed[] = []
  const boxes = new Map<string, { x: number; y: number }>()
  const slotY: Record<MindMapHalf, number> = { left: padding, right: padding }
  const leafCount: Record<MindMapHalf, number> = { left: 0, right: 0 }
  const leafTop: number[] = []
  const maxDepth: Record<MindMapHalf, number> = { left: 0, right: 0 }

  const place = (node: MindMapTreeNode, id: string, depth: number, half: MindMapHalf): number => {
    const kids = node.children ?? []
    const isCollapsed = collapsed.has(id)
    const visibleKids = isCollapsed ? [] : kids
    let y: number
    if (visibleKids.length === 0) {
      y = slotY[half]
      slotY[half] += H + siblingGap
      leafCount[half] += 1
      leafTop.push(y)
    } else {
      const childYs = visibleKids.map((kid, i) => place(kid, `${id}.${i}`, depth + 1, half))
      y = (childYs[0] + childYs[childYs.length - 1]) / 2
    }
    maxDepth[half] = Math.max(maxDepth[half], depth)
    placed.push({ node, id, depth, half, y })
    return y
  }

  for (const branch of right) place(branch.node, branch.id, 1, 'right')
  for (const branch of left) place(branch.node, branch.id, 1, 'left')

  // ---- 3. the root sits in the middle, centred on the outermost leaves ----
  const leftWidth = maxDepth.left * (W + depthGap)
  const rightWidth = maxDepth.right * (W + depthGap)
  const rootX = padding + leftWidth
  const rootY = leafTop.length === 0 ? padding : (Math.min(...leafTop) + Math.max(...leafTop)) / 2

  /** Column of a node: depth 0 is the root, depth d sits d columns out. */
  const xFor = (depth: number, half: MindMapHalf): number =>
    half === 'right'
      ? rootX + W + depthGap + (depth - 1) * (W + depthGap)
      : rootX - depthGap - W - (depth - 1) * (W + depthGap)

  boxes.set('0', { x: rootX, y: rootY })
  for (const item of placed) {
    boxes.set(item.id, { x: xFor(item.depth, item.half), y: item.y })
  }

  // ---- 4. the node boxes ----
  nodes.push({
    id: '0',
    label: root.label,
    lines: wrapLabel(root.label, labelBudget(root.block !== undefined, rootKids.length > 0)),
    block: root.block,
    depth: 0,
    x: rootX,
    y: rootY,
    hasChildren: rootKids.length > 0,
    childCount: rootKids.length,
    descendants: countMindMapNodes(root) - 1,
    collapsed: rootCollapsed,
  })
  for (const item of placed) {
    const kids = item.node.children ?? []
    nodes.push({
      id: item.id,
      label: item.node.label,
      lines: wrapLabel(
        item.node.label,
        labelBudget(item.node.block !== undefined, kids.length > 0),
      ),
      block: item.node.block,
      depth: item.depth,
      x: xFor(item.depth, item.half),
      y: item.y,
      hasChildren: kids.length > 0,
      childCount: kids.length,
      descendants: countMindMapNodes(item.node) - 1,
      collapsed: collapsed.has(item.id),
    })
  }

  // ---- 5. the connectors ----
  const dirOf = (half: MindMapHalf): 1 | -1 => (half === 'right' ? 1 : -1)
  for (const item of placed) {
    const kids = item.node.children ?? []
    if (collapsed.has(item.id)) continue
    const parent = boxes.get(item.id)!
    for (let i = 0; i < kids.length; i++) {
      const kidId = `${item.id}.${i}`
      const child = boxes.get(kidId)!
      edges.push({
        id: `${item.id}->${kidId}`,
        d: horizontalEdgePath(parent, child, dirOf(item.half)),
      })
    }
  }
  for (const branch of [...right, ...left]) {
    const child = boxes.get(branch.id)!
    edges.push({
      id: `0->${branch.id}`,
      d: horizontalEdgePath(boxes.get('0')!, child, dirOf(branch.half)),
    })
  }

  const rows = Math.max(leafCount.left, leafCount.right)
  return {
    nodes,
    edges,
    width: padding * 2 + leftWidth + W + rightWidth,
    height: rows > 0 ? rows * (H + siblingGap) - siblingGap + padding * 2 : H + padding * 2,
  }
}

/** Display-unit budget for a node label, leaving room for the badge/toggle chips. */
function labelBudget(hasBadge: boolean, hasToggle: boolean): number {
  const left = hasBadge ? 38 : 12
  const right = hasToggle ? 32 : 10
  return Math.max(6, Math.floor((MIND_MAP_METRICS.nodeWidth - left - right) / 6.5))
}

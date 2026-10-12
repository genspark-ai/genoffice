import type { AgentToolDef, MindMapTreeNode } from '@genoffice/agent-core'

/**
 * Mind map generation for the docs assistant: the tool definition the model
 * calls with the whole tree in one shot, strict-but-forgiving parsing of that
 * tree, and a pure horizontal tidy-tree layout the card component renders.
 * No runtime dependencies — layout is plain arithmetic, rendering is SVG.
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

/** Bounded recursion: the schema inlines down to the enforced depth limit; deeper
 *  levels (which the parser trims anyway) carry a generic children description. */
const nodeSchema = (depth: number): Record<string, unknown> => ({
  type: 'object',
  properties: {
    label: {
      type: 'string',
      description:
        'short node label (2-8 words) in the document\'s language, taken from its headings or key phrases',
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
    'Call it when the user asks for a mind map / 思维导图 / 脑图 / knowledge structure / overview diagram of the document or a topic. ' +
    'Pass the entire tree in this single call: the root is the document title or the focus topic, its children are the main sections, and so on. ' +
    'Labels must be short phrases in the document\'s own language and wording — never invent topics the document does not cover. ' +
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
  /** SVG path `d`: cubic from parent right edge to child left edge */
  d: string
}

export interface MindMapLayout {
  nodes: MindMapLayoutNode[]
  edges: MindMapLayoutEdge[]
  width: number
  height: number
}

function edgePath(
  parent: { x: number; y: number },
  child: { x: number; y: number },
): string {
  const px = parent.x + MIND_MAP_METRICS.nodeWidth
  const py = parent.y + MIND_MAP_METRICS.nodeHeight / 2
  const cx = child.x
  const cy = child.y + MIND_MAP_METRICS.nodeHeight / 2
  const k = (cx - px) / 2
  return `M ${px} ${py} C ${px + k} ${py}, ${cx - k} ${cy}, ${cx} ${cy}`
}

/**
 * Horizontal tidy tree: leaves take sequential x slots, parents center over
 * their children, depth maps to y. Collapsed subtrees are skipped entirely
 * (the node keeps a +N badge). Pure — same inputs always lay out the same.
 */
export function layoutMindMap(
  root: MindMapTreeNode,
  collapsed: ReadonlySet<string> = new Set(),
): MindMapLayout {
  const nodes: MindMapLayoutNode[] = []
  const edges: MindMapLayoutEdge[] = []
  let nextX = 0
  let maxDepth = 0

  const place = (node: MindMapTreeNode, id: string, depth: number): MindMapLayoutNode => {
    maxDepth = Math.max(maxDepth, depth)
    const kids = node.children ?? []
    const isCollapsed = collapsed.has(id)
    const placedKids = isCollapsed
      ? []
      : kids.map((kid, i) => place(kid, `${id}.${i}`, depth + 1))
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
      y: depth * (MIND_MAP_METRICS.nodeHeight + MIND_MAP_METRICS.depthGap),
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
  const { nodeWidth: w, nodeHeight: h, depthGap, siblingGap, padding } = MIND_MAP_METRICS
  return {
    nodes,
    edges,
    width: Math.max(w, nextX - siblingGap) + padding * 2,
    height: (maxDepth + 1) * (h + depthGap) - depthGap + padding * 2,
  }
}

/** Display-unit budget for a node label, leaving room for the badge/toggle chips. */
function labelBudget(hasBadge: boolean, hasToggle: boolean): number {
  const left = hasBadge ? 38 : 12
  const right = hasToggle ? 32 : 10
  return Math.max(6, Math.floor((MIND_MAP_METRICS.nodeWidth - left - right) / 6.5))
}

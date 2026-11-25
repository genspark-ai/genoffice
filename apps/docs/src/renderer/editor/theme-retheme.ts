import type { Editor } from '@tiptap/core'
import type { Mark, Node as PMNode } from '@tiptap/pm/model'
import type { Transaction } from '@tiptap/pm/state'
import type { Block } from '@genoffice/docx-engine'
import { TRACK_IGNORE } from './revisions'

/**
 * The page holds theme colours and faces resolved when the file was read: table
 * shading and borders, shape fills and lines, chart palettes, the face of a run
 * on a theme font. A theme pick re-reads the file under the picked theme; what
 * differs between the two readings is what the theme decides, and it is put on
 * the live document wherever the old reading still stands. Anything the user
 * changed no longer holds the old value and is left alone.
 */

/** one reading of the file: its blocks and the editor document they make */
export interface BodyReading {
  blocks: Block[]
  doc: PMNode
}

type Path = (string | number)[]

interface Change {
  path: Path
  from: unknown
  to: unknown
}

/** the values one attribute path of one node type goes to, by the value it has */
type ValueRules = Map<string, unknown>

interface BlockPatch {
  /** node or mark type + attribute path -> value -> new value, where the whole block agrees */
  byValue: Map<string, { path: Path; rules: ValueRules }>
  /** where it does not agree: place in the block + attribute path -> the one change there */
  byPlace: Map<string, Change[]>
}

/** what a theme changes in the body, per top-level block of the file */
export type ThemePatch = Map<number, BlockPatch>

const isPlain = (v: unknown): v is Record<string, unknown> =>
  typeof v === 'object' && v !== null && !Array.isArray(v)

const key = (v: unknown): string => JSON.stringify(v) ?? 'undefined'

function leafChanges(a: unknown, b: unknown, path: Path, out: Change[]): void {
  if (a === b) return
  if (isPlain(a) && isPlain(b)) {
    for (const k of new Set([...Object.keys(a), ...Object.keys(b)]))
      leafChanges(a[k], b[k], [...path, k], out)
    return
  }
  if (Array.isArray(a) && Array.isArray(b) && a.length === b.length) {
    a.forEach((item, i) => leafChanges(item, b[i], [...path, i], out))
    return
  }
  if (key(a) !== key(b)) out.push({ path, from: a, to: b })
}

const valueAt = (root: unknown, path: Path): unknown =>
  path.reduce<unknown>(
    (v, k) => (v === null || typeof v !== 'object' ? undefined : (v as never)[k]),
    root,
  )

function withValueAt(root: unknown, path: Path, value: unknown): unknown {
  if (path.length === 0) return value
  const [head, ...rest] = path
  if (Array.isArray(root)) {
    const next = root.slice()
    next[head as number] = withValueAt(root[head as number], rest, value)
    return next
  }
  const obj = (isPlain(root) ? root : {}) as Record<string, unknown>
  return { ...obj, [head as string]: withValueAt(obj[head as string], rest, value) }
}

/** children by type and count among their kind: typing in a paragraph does not renumber its drawings */
function eachChild(node: PMNode, place: string, visit: (child: PMNode, place: string) => void) {
  const seen = new Map<string, number>()
  node.forEach((child) => {
    if (child.isText) return
    const n = seen.get(child.type.name) ?? 0
    seen.set(child.type.name, n + 1)
    visit(child, `${place}/${child.type.name}:${n}`)
  })
}

// a mark has no place of its own in a paragraph that was typed in: it is known by what it held
const markType = (mark: Mark): string => `mark:${mark.type.name}`
const markPlace = (mark: Mark): string => `${markType(mark)}=${key(mark.attrs)}`

/** one node or mark of the file's reading, and what the picked theme makes of it */
interface Item {
  type: string
  place: string
  attrs: unknown
  themed?: unknown
}

function itemsOf(from: PMNode, to: PMNode): Item[] | null {
  const items: Item[] = []
  let comparable = true
  const walk = (a: PMNode, b: PMNode, place: string) => {
    if (a.type !== b.type) {
      comparable = false
      return
    }
    items.push({ type: a.type.name, place, attrs: a.attrs, themed: b.attrs })
    if (a.isTextblock) {
      const spans: Array<{ from: number; to: number; marks: readonly Mark[] }> = []
      b.forEach((child, offset) => {
        if (child.isText)
          spans.push({ from: offset, to: offset + child.nodeSize, marks: child.marks })
      })
      a.forEach((child, offset) => {
        if (!child.isText) return
        const other = spans.find((s) => s.from <= offset && s.to > offset)
        for (const mark of child.marks) {
          const now = other && mark.type.isInSet(other.marks)
          items.push({
            type: markType(mark),
            place: markPlace(mark),
            attrs: mark.attrs,
            ...(now ? { themed: now.attrs } : {}),
          })
        }
      })
    }
    const others = new Map<string, PMNode>()
    eachChild(b, place, (child, at) => others.set(at, child))
    eachChild(a, place, (child, at) => {
      const other = others.get(at)
      if (other) walk(child, other, at)
      else comparable = false
    })
  }
  walk(from, to, '')
  return comparable ? items : null
}

function blockPatch(from: PMNode, to: PMNode): BlockPatch | null {
  const items = itemsOf(from, to)
  if (!items) return null
  const changed: Array<{ item: Item; change: Change }> = []
  for (const item of items) {
    if (item.themed === undefined) continue
    const changes: Change[] = []
    leafChanges(item.attrs, item.themed, [], changes)
    for (const change of changes) changed.push({ item, change })
  }
  if (changed.length === 0) return null

  // a value that goes two ways within the block (a literal that happens to equal a
  // themed one) cannot be told by value: those are matched where they sit
  const targets = new Map<string, { path: Path; byValue: Map<string, Map<string, unknown>> }>()
  for (const { item, change } of changed) {
    const id = `${item.type}|${key(change.path)}`
    if (!targets.has(id)) targets.set(id, { path: change.path, byValue: new Map() })
  }
  for (const item of items) {
    for (const [id, { path, byValue }] of targets) {
      if (!id.startsWith(`${item.type}|`)) continue
      const was = valueAt(item.attrs, path)
      const now = item.themed === undefined ? was : valueAt(item.themed, path)
      const seen = byValue.get(key(was)) ?? new Map<string, unknown>()
      byValue.set(key(was), seen)
      seen.set(key(now), now)
    }
  }
  const patch: BlockPatch = { byValue: new Map(), byPlace: new Map() }
  for (const [id, { path, byValue }] of targets) {
    const rules: ValueRules = new Map()
    for (const [was, nows] of byValue) {
      if (nows.size !== 1) continue
      const [[nowKey, now]] = [...nows] as [[string, unknown]]
      if (nowKey !== was) rules.set(was, now)
    }
    if (rules.size > 0) patch.byValue.set(id, { path, rules })
  }
  for (const { item, change } of changed) {
    const id = `${item.type}|${key(change.path)}`
    if (patch.byValue.get(id)?.rules.has(key(change.from))) continue
    const there = patch.byPlace.get(item.place) ?? []
    if (!there.some((c) => key(c.path) === key(change.path))) there.push(change)
    patch.byPlace.set(item.place, there)
  }
  return patch
}

/** `attrs` of a node or mark of `type` at `place` with the block's changes put on; the same object when none applies */
function themedAttrs(block: BlockPatch, type: string, place: string, attrs: unknown): unknown {
  let next = attrs
  for (const [id, { path, rules }] of block.byValue) {
    if (!id.startsWith(`${type}|`)) continue
    const was = key(valueAt(next, path))
    if (rules.has(was)) next = withValueAt(next, path, rules.get(was))
  }
  for (const change of block.byPlace.get(place) ?? []) {
    if (key(valueAt(next, change.path)) === key(change.from))
      next = withValueAt(next, change.path, change.to)
  }
  return next
}

const blockIndex = (node: PMNode): number | null => {
  const idx: unknown = node.attrs.docxIndex
  return typeof idx === 'number' ? idx : null
}

/** what differs between two readings of one file, `from` under the theme on the page and `to` under the picked one */
export function themePatch(from: PMNode, to: PMNode): ThemePatch {
  const patch: ThemePatch = new Map()
  const targets = new Map<number, PMNode>()
  to.forEach((node) => {
    const idx = blockIndex(node)
    if (idx !== null) targets.set(idx, node)
  })
  from.forEach((node) => {
    const idx = blockIndex(node)
    const target = idx === null ? undefined : targets.get(idx)
    const block = target && blockPatch(node, target)
    if (block) patch.set(idx!, block)
  })
  return patch
}

/** the patch as steps on `tr`; true when anything on the page changed */
export function applyThemePatch(tr: Transaction, patch: ThemePatch): boolean {
  const doc = tr.doc
  let touched = false
  const visit = (node: PMNode, pos: number, place: string, block: BlockPatch) => {
    const attrs = themedAttrs(block, node.type.name, place, node.attrs)
    if (attrs !== node.attrs) {
      tr.setNodeMarkup(pos, undefined, attrs as PMNode['attrs'])
      touched = true
    }
    if (node.isTextblock) {
      node.forEach((child, offset) => {
        if (!child.isText) return
        for (const mark of child.marks) {
          const next = themedAttrs(block, markType(mark), markPlace(mark), mark.attrs)
          if (next === mark.attrs) continue
          const at = pos + 1 + offset
          tr.removeMark(at, at + child.nodeSize, mark)
          tr.addMark(at, at + child.nodeSize, mark.type.create(next as Mark['attrs']))
          touched = true
        }
      })
    }
    const seen = new Map<string, number>()
    node.forEach((child, offset) => {
      if (child.isText) return
      const n = seen.get(child.type.name) ?? 0
      seen.set(child.type.name, n + 1)
      visit(child, pos + 1 + offset, `${place}/${child.type.name}:${n}`, block)
    })
  }
  doc.forEach((node, offset) => {
    const idx = blockIndex(node)
    const block = idx === null ? undefined : patch.get(idx)
    if (block) visit(node, offset, '', block)
  })
  return touched
}

/** repaints the page for a theme pick; the change is the theme's, not an edit of its own */
export function rethemeBody(editor: Editor, from: PMNode, to: PMNode): void {
  const patch = themePatch(from, to)
  if (patch.size === 0) return
  const tr = editor.state.tr
  if (!applyThemePatch(tr, patch)) return
  editor.view.dispatch(tr.setMeta(TRACK_IGNORE, true).setMeta('addToHistory', false))
}

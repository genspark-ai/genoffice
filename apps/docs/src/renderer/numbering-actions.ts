/**
 * List numbering actions: allocating numIds, new list definitions, restart /
 * continue numbering. Extracted from App.tsx; the App component passes a
 * NumberingContext built fresh per call so state never goes stale.
 */
import type { Editor } from '@tiptap/core'
import type { Transaction } from '@tiptap/pm/state'
import {
  computeListValues,
  customLevelFromNumberingLevel,
  numberingLevelFromCustom,
  type CustomNumberingLevel,
  type ListItemRef,
  type NumberingDef,
  type NumberingLevel,
} from '@genoffice/docx-engine'
import type { DocState, PendingNumbering } from './doc-state'
import { t } from './i18n/locale'

/** The App state the numbering actions need; built fresh per call. */
export interface NumberingContext {
  editor: Editor | null
  doc: DocState | null
  pendingNumbering: PendingNumbering
  setPendingNumbering: (update: (prev: PendingNumbering) => PendingNumbering) => void
  numIdFloorRef: { current: number }
  setStatus: (status: string) => void
  /** a level linked to a paragraph style also puts w:numPr on that style (Word writes both) */
  linkStyle?: (styleId: string, numId: string, ilvl: number) => void
}

export function nextNumId(ctx: NumberingContext): string {
  let max = 2 // the blank template occupies 1/2
  for (const id of ctx.doc?.parsed.numbering.keys() ?? [])
    max = Math.max(max, parseInt(id, 10) || 0)
  for (const d of ctx.pendingNumbering.newDefs) max = Math.max(max, parseInt(d.numId, 10) || 0)
  for (const r of ctx.pendingNumbering.restartNums) max = Math.max(max, parseInt(r.numId, 10) || 0)
  max = Math.max(max, ctx.numIdFloorRef.current)
  ctx.numIdFloorRef.current = max + 1
  return String(max + 1)
}

/** Instant display of editor markers: inject pending definitions into listNumbering storage (re-parsing takes over after save) */
export function overlayNumberingDef(ctx: NumberingContext, def: NumberingDef): void {
  if (!ctx.editor) return
  const store = ctx.editor.storage.listNumbering as { defs: Map<string, NumberingDef> }
  store.defs = new Map(store.defs).set(def.numId, def)
}

/** Fallback when a new list can't reuse a numId: adopt an existing same-kind definition, otherwise create a new one */
export function createNumberingDef(ctx: NumberingContext, kind: 'bullet' | 'ordered'): string {
  // brand-new abstractNum + num (blank template style; when the part is missing the engine creates the part too)
  const numId = nextNumId(ctx)
  ctx.setPendingNumbering((p) => ({ ...p, newDefs: [...p.newDefs, { numId, kind }] }))
  overlayNumberingDef(ctx, {
    numId,
    abstractNumId: `pending-${numId}`,
    levels: Object.fromEntries(
      Array.from({ length: 5 }, (_, i) => [
        i,
        {
          numFmt: kind === 'bullet' ? 'bullet' : 'decimal',
          lvlText: kind === 'bullet' ? '' : `%${i + 1}.`,
          start: 1,
          indentLeft: 720 * (i + 1),
          hanging: 360,
        },
      ]),
    ),
    startOverrides: {},
  })
  return numId
}

/** New list definitions with custom levels (bullet library / numbering library / multilevel gallery / define dialog) */
export function createCustomListDef(
  ctx: NumberingContext,
  levels: CustomNumberingLevel[],
): string | null {
  if (!ctx.doc || levels.length === 0) return null
  const kind = levels[0].numFmt === 'bullet' ? ('bullet' as const) : ('ordered' as const)
  const numId = nextNumId(ctx)
  ctx.setPendingNumbering((p) => ({ ...p, newDefs: [...p.newDefs, { numId, kind, levels }] }))
  overlayNumberingDef(ctx, {
    numId,
    abstractNumId: `pending-${numId}`,
    levels: Object.fromEntries(
      levels.map((l, i) => [i, numberingLevelFromCustom(l, picBulletSrc(ctx, l.picBulletId))]),
    ),
    startOverrides: {},
  })
  levels.forEach((l, i) => {
    if (l.pStyle) ctx.linkStyle?.(l.pStyle, numId, i)
  })
  return numId
}

/** data URL of a pending picture bullet (parsed ones already carry picBulletSrc on the level) */
function picBulletSrc(ctx: NumberingContext, id: number | undefined): string | undefined {
  if (id === undefined) return undefined
  const pic = ctx.pendingNumbering.picBullets.find((p) => p.id === id)
  if (pic) return `data:${pic.mime};base64,${pic.base64}`
  for (const def of ctx.doc?.parsed.numbering.values() ?? [])
    for (const level of Object.values(def.levels))
      if (level.picBulletId === id && level.picBulletSrc) return level.picBulletSrc
  return undefined
}

/** Register picture-bullet bytes; the returned id goes into a level's picBulletId */
export function addPictureBullet(
  ctx: NumberingContext,
  image: { base64: string; mime: 'image/png' | 'image/jpeg' | 'image/gif' },
): number {
  let max = 0
  for (const def of ctx.doc?.parsed.numbering.values() ?? [])
    for (const level of Object.values(def.levels))
      if (level.picBulletId !== undefined) max = Math.max(max, level.picBulletId)
  for (const pic of ctx.pendingNumbering.picBullets) max = Math.max(max, pic.id)
  const id = max + 1
  ctx.setPendingNumbering((p) => ({ ...p, picBullets: [...p.picBullets, { id, ...image }] }))
  // the overlay for a definition created in the same click reads the id from this list
  ctx.pendingNumbering = {
    ...ctx.pendingNumbering,
    picBullets: [...ctx.pendingNumbering.picBullets, { id, ...image }],
  }
  return id
}

/** The list the caret is in: its definition (overlay first) and the caret's level */
export function currentListTarget(ctx: NumberingContext): {
  numId: string
  ilvl: number
  def: NumberingDef
  level: NumberingLevel | undefined
} | null {
  if (!ctx.editor || !ctx.doc) return null
  const attrs = ctx.editor.getAttributes('docListItem')
  const numId = attrs?.numId as string | null | undefined
  if (numId == null) return null
  const store = ctx.editor.storage.listNumbering as { defs: Map<string, NumberingDef> }
  const def = store.defs.get(String(numId)) ?? ctx.doc.parsed.numbering.get(String(numId))
  if (!def) return null
  const ilvl = Number(attrs.ilvl) || 0
  return { numId: String(numId), ilvl, def, level: def.levels[ilvl] }
}

/**
 * Rewrite one level of the caret's list definition (Adjust List Indents, Define on
 * an existing list). Word edits the abstractNum in place, so every list sharing
 * it changes: pending definitions are patched in the pending state, parsed ones
 * become a levelEdit for the save.
 */
export function editListLevel(
  ctx: NumberingContext,
  patch: Partial<CustomNumberingLevel>,
  ilvlOverride?: number,
): boolean {
  const target = currentListTarget(ctx)
  if (!target || !ctx.editor) return false
  const ilvl = ilvlOverride ?? target.ilvl
  const base: CustomNumberingLevel = target.def.levels[ilvl]
    ? customLevelFromNumberingLevel(target.def.levels[ilvl])
    : { numFmt: 'decimal', lvlText: `%${ilvl + 1}.`, indentLeft: 720 * (ilvl + 1), hanging: 360 }
  const level: CustomNumberingLevel = { ...base, ...patch }
  const abstractNumId = target.def.abstractNumId
  if (abstractNumId.startsWith('pending-')) {
    const pendingNumId = abstractNumId.slice('pending-'.length)
    ctx.setPendingNumbering((p) => ({
      ...p,
      newDefs: p.newDefs.map((d) => {
        if (d.numId !== pendingNumId) return d
        const levels = d.levels
          ? [...d.levels]
          : Object.keys(target.def.levels)
              .map(Number)
              .sort((a, b) => a - b)
              .map((i) => customLevelFromNumberingLevel(target.def.levels[i]))
        while (levels.length <= ilvl)
          levels.push({
            numFmt: 'decimal',
            lvlText: `%${levels.length + 1}.`,
            indentLeft: 720 * (levels.length + 1),
            hanging: 360,
          })
        levels[ilvl] = level
        return { ...d, levels }
      }),
    }))
  } else {
    ctx.setPendingNumbering((p) => ({
      ...p,
      levelEdits: [
        ...p.levelEdits.filter((e) => !(e.abstractNumId === abstractNumId && e.ilvl === ilvl)),
        { abstractNumId, ilvl, level },
      ],
    }))
  }
  const store = ctx.editor.storage.listNumbering as { defs: Map<string, NumberingDef> }
  const overlay = numberingLevelFromCustom(level, picBulletSrc(ctx, level.picBulletId))
  const next = new Map(store.defs)
  for (const [id, def] of next) {
    if (def.abstractNumId !== abstractNumId) continue
    next.set(id, { ...def, levels: { ...def.levels, [ilvl]: overlay } })
  }
  store.defs = next
  if (level.pStyle) ctx.linkStyle?.(level.pStyle, target.numId, ilvl)
  // the marker plugin recomputes when the defs table identity changes on the next transaction
  ctx.editor.view.dispatch(ctx.editor.state.tr)
  return true
}

/** The number the caret's item shows now (1 for bullets / unknown definitions) */
export function currentListValue(editor: Editor | null, numId: string, ilvl: number): number {
  if (!editor) return 1
  const store = editor.storage.listNumbering as { defs: Map<string, NumberingDef> }
  const refs: ListItemRef[] = []
  let caretIndex = -1
  const { $from } = editor.state.selection
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name !== 'docListItem') return true
    if (pos <= $from.pos && $from.pos <= pos + node.nodeSize) caretIndex = refs.length
    refs.push({
      numId: (node.attrs.numId as string | null) ?? null,
      ilvl: Number(node.attrs.ilvl) || 0,
    })
    return false
  })
  if (caretIndex < 0) return 1
  const values = computeListValues(refs, store.defs)
  void numId
  void ilvl
  return values[caretIndex] ?? 1
}

/** whether some earlier list (a different numId) exists for "Continue from previous list" */
export function hasPreviousList(editor: Editor | null, numId: string): boolean {
  if (!editor) return false
  const startIdx = editor.state.selection.$from.index(0)
  let found = false
  editor.state.doc.forEach((node, _offset, idx) => {
    if (idx >= startIdx || found) return
    if (node.type.name === 'docListItem' && node.attrs.numId && node.attrs.numId !== numId)
      found = true
  })
  return found
}

/** Right-click / gallery: move the caret's list items to a level (1-9 shown to the user as 0-8 here) */
export function changeListLevel(editor: Editor, ilvl: number): boolean {
  if (!editor.isActive('docListItem')) return false
  const next = Math.min(Math.max(ilvl, 0), 8)
  return editor.chain().focus().updateAttributes('docListItem', { ilvl: next }).run()
}

/**
 * Set Numbering Value: start a new list at `value` (new w:num on the same
 * abstractNum with a startOverride), or continue the previous list — with a
 * skipped-to value when the user advances it.
 */
export function setNumberingValue(
  ctx: NumberingContext,
  opts: { mode: 'restart' | 'continue'; value?: number },
): boolean {
  if (!ctx.editor || !ctx.doc) return false
  const target = currentListTarget(ctx)
  if (!target) return false
  const { state, view } = ctx.editor
  const startIdx = state.selection.$from.index(0)
  const store = ctx.editor.storage.listNumbering as { defs: Map<string, NumberingDef> }
  // one transaction whether the item merges into the previous list, skips to a value, or both
  let tr = state.tr
  let sourceNumId = target.numId
  let sourceDef = target.def
  if (opts.mode === 'continue') {
    const prevNumId = previousListNumId(ctx.editor, target.numId)
    if (!prevNumId) return false
    tr = rewriteNumIdForwardSteps(tr, startIdx, target.numId, prevNumId)
    sourceNumId = prevNumId
    sourceDef = store.defs.get(prevNumId) ?? ctx.doc.parsed.numbering.get(prevNumId) ?? target.def
    if (opts.value === undefined) {
      view.dispatch(tr)
      ctx.setStatus(t('appNumberingContinued'))
      return true
    }
  }
  const value = Math.max(0, Math.floor(opts.value ?? 1))
  const numId = nextNumId(ctx)
  ctx.setPendingNumbering((p) => ({
    ...p,
    restartNums: [
      ...p.restartNums,
      { numId, abstractNumId: sourceDef.abstractNumId, startOverrides: { [target.ilvl]: value } },
    ],
  }))
  overlayNumberingDef(ctx, { ...sourceDef, numId, startOverrides: { [target.ilvl]: value } })
  view.dispatch(rewriteNumIdForwardSteps(tr, startIdx, sourceNumId, numId))
  ctx.setStatus(t('appNumberingValueSet', { value: String(value) }))
  return true
}

export function allocateListNumId(
  ctx: NumberingContext,
  kind: 'bullet' | 'ordered',
): string | null {
  if (!ctx.doc) return null
  // the document already has a same-kind numbering definition (even if unused in the body): the new num points at its abstractNum
  const match = [...ctx.doc.parsed.numbering.values()].find(
    (d) => (d.levels[0]?.numFmt === 'bullet') === (kind === 'bullet'),
  )
  if (match) {
    const numId = nextNumId(ctx)
    ctx.setPendingNumbering((p) => ({
      ...p,
      restartNums: [
        ...p.restartNums,
        { numId, abstractNumId: match.abstractNumId, startOverrides: { 0: 1 } },
      ],
    }))
    overlayNumberingDef(ctx, { ...match, numId, startOverrides: { 0: 1 } })
    return numId
  }
  return createNumberingDef(ctx, kind)
}

interface PmJsonNode {
  type?: string
  attrs?: Record<string, unknown>
  content?: PmJsonNode[]
}

/**
 * Pending definitions no list item references (a list made by the ribbon or
 * autocorrect, then undone or reverted with Backspace) are left out of the
 * save. Undo/redo only touch the document, so the document is the one source
 * of truth for which definitions matter.
 */
export function pruneUnreferencedNumbering(
  pending: PendingNumbering,
  doc: PmJsonNode,
): PendingNumbering {
  const used = new Set<string>()
  const walk = (node: PmJsonNode) => {
    if (node.type === 'docListItem' && node.attrs?.numId != null) used.add(String(node.attrs.numId))
    node.content?.forEach(walk)
  }
  walk(doc)
  const newDefs = pending.newDefs.filter((d) => used.has(d.numId))
  const picIds = new Set<number>()
  for (const d of newDefs)
    for (const l of d.levels ?? []) if (l.picBulletId !== undefined) picIds.add(l.picBulletId)
  for (const e of pending.levelEdits)
    if (e.level.picBulletId !== undefined) picIds.add(e.level.picBulletId)
  return {
    newDefs,
    restartNums: pending.restartNums.filter((r) => used.has(r.numId)),
    levelEdits: pending.levelEdits,
    picBullets: pending.picBullets.filter((p) => picIds.has(p.id)),
  }
}

/** Steps for "from the cursor's block onward, items of numId become newNumId", appended to `tr` */
function rewriteNumIdForwardSteps(
  tr: Transaction,
  startIdx: number,
  oldNumId: string,
  newNumId: string,
): Transaction {
  tr.doc.forEach((node, offset, idx) => {
    if (idx < startIdx) return
    if (node.type.name === 'docListItem' && node.attrs.numId === oldNumId) {
      tr.setNodeMarkup(offset, undefined, { ...node.attrs, numId: newNumId })
    }
  })
  return tr
}

/** From the cursor's block onward, move list items sharing the numId to a new numId (Word: restart applies to all later items of that list) */
export function rewriteNumIdForward(
  ctx: NumberingContext,
  oldNumId: string,
  newNumId: string,
): void {
  if (!ctx.editor) return
  const { state, view } = ctx.editor
  view.dispatch(
    rewriteNumIdForwardSteps(state.tr, state.selection.$from.index(0), oldNumId, newNumId),
  )
}

/** numId of the nearest earlier list that is not the caret's own */
function previousListNumId(editor: Editor, curNumId: string): string | null {
  const startIdx = editor.state.selection.$from.index(0)
  let prev: string | null = null
  editor.state.doc.forEach((node, _offset, idx) => {
    if (idx >= startIdx) return
    if (node.type.name === 'docListItem' && node.attrs.numId && node.attrs.numId !== curNumId)
      prev = node.attrs.numId as string
  })
  return prev
}

/** Context menu: restart numbering (new num + startOverride pointing at the same abstractNum;
 *  lists with no definition (numId=0 / CSS counters) get a new definition, naturally starting at 1) */
export function restartNumbering(ctx: NumberingContext): void {
  if (!ctx.editor || !ctx.doc) return
  const attrs = ctx.editor.getAttributes('docListItem')
  const oldNumId = attrs?.numId as string | null
  if (oldNumId == null) return
  const store = ctx.editor.storage.listNumbering as { defs: Map<string, NumberingDef> }
  const def = store.defs.get(String(oldNumId)) ?? ctx.doc.parsed.numbering.get(String(oldNumId))
  let numId: string
  if (def) {
    const ilvl = Number(attrs.ilvl) || 0
    numId = nextNumId(ctx)
    ctx.setPendingNumbering((p) => ({
      ...p,
      restartNums: [
        ...p.restartNums,
        { numId, abstractNumId: def.abstractNumId, startOverrides: { [ilvl]: 1 } },
      ],
    }))
    overlayNumberingDef(ctx, { ...def, numId, startOverrides: { [ilvl]: 1 } })
  } else {
    numId = createNumberingDef(ctx, (attrs.kind as 'bullet' | 'ordered') ?? 'ordered')
  }
  rewriteNumIdForward(ctx, String(oldNumId), numId)
  ctx.setStatus(t('appNumberingRestarted'))
}

/** Context menu: continue numbering (merge back into the previous list's numId; the count continues) */
export function continueNumbering(ctx: NumberingContext): void {
  if (!ctx.editor) return
  const attrs = ctx.editor.getAttributes('docListItem')
  const curNumId = attrs?.numId as string | null
  if (!curNumId) return
  const prevNumId = previousListNumId(ctx.editor, String(curNumId))
  if (!prevNumId) return
  rewriteNumIdForward(ctx, String(curNumId), prevNumId)
  ctx.setStatus(t('appNumberingContinued'))
}

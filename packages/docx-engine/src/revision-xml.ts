/**
 * Accept / reject tracked changes in a raw WordprocessingML part (headers,
 * footers, footnotes, endnotes): the parts the editor never loads as nodes.
 * Word's Accept All semantics on the markup itself: deleted runs leave,
 * inserted runs stay unwrapped, moved text stays only at its destination,
 * property-change records drop, a deleted paragraph mark joins the paragraph
 * into the next one, a deleted row goes.
 */

export type RevisionXmlMode = 'accept' | 'reject'

export type RevisionXmlType = 'insertion' | 'deletion' | 'formatting' | 'move'

export interface RevisionXmlFilter {
  /** case-insensitive w:author match */
  author?: string
  type?: RevisionXmlType
  /** ISO date: only changes recorded strictly before it */
  before?: string
  /** exact w:date match */
  date?: string
  /** w:id of the wrapper / record */
  ids?: string[]
}

/** one accept / reject the editor applied to a part it holds as a display model */
export interface RevisionResolution {
  mode: RevisionXmlMode
  /** absent = every tracked change in the part */
  filter?: RevisionXmlFilter
}

export interface RevisionXmlResult {
  xml: string
  /** revisions the mode was applied to */
  applied: number
  /** revisions still in the part afterwards */
  remaining: number
}

interface El {
  name: string
  attrs: string
  start: number
  end: number
  innerStart: number
  innerEnd: number
  selfClosing: boolean
  parent: number
}

const TAG_RE =
  /<(\/?)([A-Za-z][\w.-]*:?[\w.-]*)((?:\s+[^>]*?)?)(\/?)>|<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>/g

function scan(xml: string): El[] {
  const out: El[] = []
  const stack: number[] = []
  TAG_RE.lastIndex = 0
  for (let m = TAG_RE.exec(xml); m; m = TAG_RE.exec(xml)) {
    if (m[2] === undefined) continue
    const [whole, close, name, attrs, self] = m
    const start = m.index
    const end = start + whole.length
    if (close) {
      for (let i = stack.length - 1; i >= 0; i--) {
        const el = out[stack[i]!]!
        if (el.name === name) {
          el.innerEnd = start
          el.end = end
          stack.length = i
          break
        }
      }
      continue
    }
    const el: El = {
      name,
      attrs: attrs ?? '',
      start,
      end,
      innerStart: end,
      innerEnd: end,
      selfClosing: !!self,
      parent: stack.length ? stack[stack.length - 1]! : -1,
    }
    out.push(el)
    if (!self) stack.push(out.length - 1)
  }
  return out
}

const attrOf = (el: El, name: string): string | undefined =>
  new RegExp(`\\s${name}="([^"]*)"`).exec(el.attrs)?.[1]

const PROP_CHANGE: Record<string, string> = {
  'w:rPrChange': 'w:rPr',
  'w:pPrChange': 'w:pPr',
  'w:sectPrChange': 'w:sectPr',
  'w:tblPrChange': 'w:tblPr',
  'w:tblPrExChange': 'w:tblPrEx',
  'w:trPrChange': 'w:trPr',
  'w:tcPrChange': 'w:tcPr',
  'w:tblGridChange': 'w:tblGrid',
}

const RANGE_MARKERS = new Set([
  'w:moveFromRangeStart',
  'w:moveFromRangeEnd',
  'w:moveToRangeStart',
  'w:moveToRangeEnd',
])

/** what a revision element records; undefined = not a revision */
function typeOf(el: El): RevisionXmlType | undefined {
  switch (el.name) {
    case 'w:ins':
    case 'w:cellIns':
      return 'insertion'
    case 'w:del':
    case 'w:cellDel':
      return 'deletion'
    case 'w:moveFrom':
    case 'w:moveTo':
      return 'move'
    case 'w:numberingChange':
    case 'w:cellMerge':
      return 'formatting'
    default:
      if (PROP_CHANGE[el.name]) return 'formatting'
      return undefined
  }
}

function matches(el: El, filter: RevisionXmlFilter | undefined, type: RevisionXmlType): boolean {
  if (!filter) return true
  if (filter.type !== undefined && filter.type !== type) return false
  if (filter.author !== undefined) {
    const author = attrOf(el, 'w:author') ?? ''
    if (author.toLowerCase() !== filter.author.toLowerCase()) return false
  }
  if (filter.before !== undefined) {
    const date = attrOf(el, 'w:date')
    if (!date || Date.parse(date) >= Date.parse(filter.before)) return false
  }
  if (filter.date !== undefined && (attrOf(el, 'w:date') ?? '') !== filter.date) return false
  if (filter.ids !== undefined && !filter.ids.includes(attrOf(el, 'w:id') ?? '')) return false
  return true
}

interface Edit {
  start: number
  end: number
  text: string
}

const parentOf = (els: El[], el: El): El | undefined =>
  el.parent >= 0 ? els[el.parent] : undefined

/** `w:ins` / `w:del` carry the paragraph-mark and row-level records as empty elements */
function markerSlot(els: El[], el: El): 'paraMark' | 'row' | 'other' | null {
  if (!el.selfClosing) return null
  const parent = parentOf(els, el)
  if (!parent) return 'other'
  if (parent.name === 'w:trPr') return 'row'
  if (parent.name === 'w:rPr' && parentOf(els, parent)?.name === 'w:pPr') return 'paraMark'
  return 'other'
}

const ancestor = (els: El[], el: El, name: string): El | undefined => {
  for (let p = parentOf(els, el); p; p = parentOf(els, p)) if (p.name === name) return p
  return undefined
}

const childOf = (els: El[], el: El, name: string): El | undefined =>
  els.find((c) => c.parent === els.indexOf(el) && c.name === name)

function children(els: El[], el: El): El[] {
  const idx = els.indexOf(el)
  return els.filter((c) => c.parent === idx)
}

/** the paragraph right after `p` in the same parent (markers like bookmarkEnd may sit between) */
function nextParagraph(els: El[], p: El): El | undefined {
  const idx = els.indexOf(p)
  for (let i = idx + 1; i < els.length; i++) {
    const c = els[i]!
    if (c.start < p.end) continue
    if (c.parent !== p.parent) return undefined
    if (c.name === 'w:p') return c
    if (!/^w:(bookmark|commentRange|permStart|permEnd|proofErr)/.test(c.name)) return undefined
  }
  return undefined
}

/** join `p` into the following paragraph: `p` keeps its properties, the next paragraph's open tag and pPr go */
function mergeEdits(els: El[], p: El): Edit[] | null {
  const next = nextParagraph(els, p)
  if (!next || next.selfClosing) return null
  const nextPPr = childOf(els, next, 'w:pPr')
  const closeStart = p.innerEnd
  return [
    { start: closeStart, end: p.end, text: '' },
    { start: next.start, end: nextPPr ? nextPPr.end : next.innerStart, text: '' },
  ]
}

function unwrap(el: El): Edit[] {
  return [
    { start: el.start, end: el.innerStart, text: '' },
    { start: el.innerEnd, end: el.end, text: '' },
  ]
}

const remove = (el: El): Edit => ({ start: el.start, end: el.end, text: '' })

/** drop an empty marker element; an rPr / trPr / pPr it leaves empty goes with it */
function removeMarker(els: El[], el: El): Edit {
  let target = el
  for (let parent = parentOf(els, target); parent; parent = parentOf(els, target)) {
    if (!/^w:(?:rPr|trPr|pPr)$/.test(parent.name) || children(els, parent).length > 1) break
    target = parent
  }
  return remove(target)
}

/** restore the pre-change properties recorded in a *Change element */
function revertPropChange(els: El[], xml: string, change: El): Edit[] {
  const holderName = PROP_CHANGE[change.name]!
  const holder = parentOf(els, change)
  if (!holder || holder.name !== holderName) return [remove(change)]
  const old = children(els, change).find((c) => c.name === holderName)
  const oldInner = old ? xml.slice(old.innerStart, old.innerEnd) : ''
  if (holderName === 'w:pPr') {
    // the paragraph mark's rPr and the sectPr are not part of the record: keep them in place
    const keepFrom = children(els, holder).find(
      (c) => c.name === 'w:rPr' || c.name === 'w:sectPr' || c === change,
    )!
    // nothing but the record: one edit, or the removal would swallow the insertion at its start
    if (keepFrom === change) return [{ start: holder.innerStart, end: change.end, text: oldInner }]
    return [
      { start: holder.innerStart, end: keepFrom.start, text: oldInner },
      { start: change.start, end: change.end, text: '' },
    ]
  }
  if (holderName === 'w:sectPr') {
    const refs = children(els, holder)
      .filter((c) => c.name === 'w:headerReference' || c.name === 'w:footerReference')
      .map((c) => xml.slice(c.start, c.end))
      .join('')
    return [{ start: holder.innerStart, end: holder.innerEnd, text: refs + oldInner }]
  }
  // the holder's own tracked markers (paragraph-mark / row ins, del, move) are
  // separate revisions: edit around them so their own removal in the same pass
  // is not swallowed by a whole-holder replacement
  const isMarker = (c: El) =>
    c.selfClosing && /^w:(?:ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge)$/.test(c.name)
  const kids = children(els, holder)
  if (kids.some(isMarker)) {
    return [
      ...kids.filter((c) => !isMarker(c)).map(remove),
      { start: holder.innerEnd, end: holder.innerEnd, text: oldInner },
    ]
  }
  if (!oldInner.trim() && holderName !== 'w:tblPr' && holderName !== 'w:tblGrid')
    return [remove(holder)]
  return [{ start: holder.innerStart, end: holder.innerEnd, text: oldInner }]
}

function planEdits(
  els: El[],
  xml: string,
  mode: RevisionXmlMode,
  filter: RevisionXmlFilter | undefined,
): { edits: Edit[]; applied: number } {
  const edits: Edit[] = []
  let applied = 0
  const act = (el: El, drop: boolean): void => {
    const slot = markerSlot(els, el)
    if (slot === 'paraMark') {
      const p = drop ? ancestor(els, el, 'w:p') : undefined
      const merged = p ? mergeEdits(els, p) : null
      edits.push(removeMarker(els, el), ...(merged ?? []))
      return
    }
    if (slot === 'row') {
      const row = ancestor(els, el, 'w:tr')
      edits.push(drop && row ? remove(row) : removeMarker(els, el))
      return
    }
    if (slot === 'other') {
      edits.push(removeMarker(els, el))
      return
    }
    if (drop) {
      edits.push(remove(el))
      return
    }
    edits.push(...unwrap(el))
    if (mode === 'reject') {
      // the deleted runs come back as live text
      const idx = els.indexOf(el)
      for (let i = idx + 1; i < els.length && els[i]!.start < el.innerEnd; i++) {
        const c = els[i]!
        if (c.name !== 'w:delText' && c.name !== 'w:delInstrText') continue
        const live = c.name === 'w:delText' ? 'w:t' : 'w:instrText'
        edits.push({ start: c.start + 1, end: c.start + 1 + c.name.length, text: live })
        if (!c.selfClosing)
          edits.push({ start: c.innerEnd + 2, end: c.innerEnd + 2 + c.name.length, text: live })
      }
    }
  }
  // the end marker of a move range carries no author / date: it goes with its start (same w:id)
  const startIds = new Set<string>()
  for (const el of els) {
    if (!RANGE_MARKERS.has(el.name)) continue
    const id = attrOf(el, 'w:id') ?? ''
    if (el.name.endsWith('Start')) {
      if (matches(el, filter, 'move')) {
        startIds.add(id)
        edits.push(remove(el))
      }
    } else if (startIds.has(id) || matches(el, filter, 'move')) edits.push(remove(el))
  }
  const goneCells = new Set<El>()
  for (const el of els) {
    if (RANGE_MARKERS.has(el.name)) continue
    const type = typeOf(el)
    if (!type || !matches(el, filter, type)) continue
    applied++
    switch (el.name) {
      case 'w:ins':
      case 'w:moveTo':
        act(el, mode === 'reject')
        break
      case 'w:del':
      case 'w:moveFrom':
        act(el, mode === 'accept')
        break
      case 'w:cellIns':
      case 'w:cellDel': {
        // a cell whose insertion is rejected, or whose deletion is accepted, leaves its row
        const gone = (el.name === 'w:cellIns') === (mode === 'reject')
        const cell = gone ? ancestor(els, el, 'w:tc') : undefined
        if (cell) goneCells.add(cell)
        else edits.push(remove(el))
        break
      }
      case 'w:numberingChange':
      case 'w:cellMerge':
        edits.push(remove(el))
        break
      default:
        edits.push(...(mode === 'accept' ? [remove(el)] : revertPropChange(els, xml, el)))
    }
  }
  // a row left without a cell goes whole: a w:tr holds at least one w:tc
  const emptied = new Set<El>()
  for (const cell of goneCells) {
    const row = parentOf(els, cell)
    const cells = row ? children(els, row).filter((c) => c.name === 'w:tc') : []
    if (row?.name === 'w:tr' && cells.every((c) => goneCells.has(c))) emptied.add(row)
    else edits.push(remove(cell))
  }
  for (const row of emptied) edits.push(remove(row))
  return { edits, applied }
}

const UNEVEN_CELL = /<w:(?:gridSpan|vMerge|hMerge)\b/
const UNEVEN_ROW = /<w:(?:gridBefore|gridAfter)\b/

/**
 * Which way inserted cells pushed the others is not in the file, but Shift
 * cells down leaves a shape of its own in a plain grid: rows added at the
 * bottom whose cells are all inserted except in the columns that moved, and
 * as many inserted cells higher up in exactly those columns, all one author's
 * at one time. The edits that move those columns back up and drop the added
 * rows, or null when the table is not that or `filter` leaves one of its
 * inserted cells pending.
 */
function downShiftEdits(
  els: El[],
  xml: string,
  at: number,
  filter: RevisionXmlFilter | undefined,
): { edits: Edit[]; applied: number } | null {
  const table = els[at]!
  if (!xml.slice(table.start, table.end).includes('<w:cellIns')) return null
  const kids = new Map<number, El[]>()
  for (let i = at + 1; i < els.length && els[i]!.start < table.end; i++) {
    const list = kids.get(els[i]!.parent)
    if (list) list.push(els[i]!)
    else kids.set(els[i]!.parent, [els[i]!])
  }
  const under = (el: El) => kids.get(els.indexOf(el, at)) ?? []
  const own = under(table)
  if (own.some((c) => c.name !== 'w:tr' && c.name !== 'w:tblPr' && c.name !== 'w:tblGrid'))
    return null
  const rows = own.filter((c) => c.name === 'w:tr')
  const grid: El[][] = []
  const records: Array<Array<El | undefined>> = []
  for (const row of rows) {
    const cells: El[] = []
    const marks: Array<El | undefined> = []
    for (const child of under(row)) {
      if (child.name === 'w:trPr') {
        if (UNEVEN_ROW.test(xml.slice(child.start, child.end))) return null
      } else if (child.name === 'w:tc') {
        const tcPr = under(child).find((c) => c.name === 'w:tcPr')
        if (tcPr && UNEVEN_CELL.test(xml.slice(tcPr.start, tcPr.end))) return null
        cells.push(child)
        marks.push(tcPr ? under(tcPr).find((c) => c.name === 'w:cellIns') : undefined)
      } else if (child.name !== 'w:tblPrEx') return null
    }
    grid.push(cells)
    records.push(marks)
  }
  const width = grid[0]?.length ?? 0
  if (grid.length < 2 || width < 2 || grid.some((row) => row.length !== width)) return null
  const pending = records.flat().filter((mark): mark is El => !!mark)
  const first = pending[0]
  if (!first) return null
  const by = (el: El) => `${attrOf(el, 'w:author') ?? ''}\n${attrOf(el, 'w:date') ?? ''}`
  if (pending.some((mark) => by(mark) !== by(first) || !matches(mark, filter, 'insertion')))
    return null
  const bottom = grid.length - 1
  const moved = records[bottom]!.flatMap((mark, col) => (mark ? [] : [col]))
  if (moved.length === 0 || moved.length === width) return null
  if (moved.some((col, i) => i > 0 && col !== moved[i - 1]! + 1)) return null
  const left = moved[0]!
  const right = left + moved.length
  const added = (row: number) =>
    records[row]!.every((mark, col) => (col >= left && col < right ? !mark : !!mark))
  let height = 0
  while (height < bottom && added(bottom - height)) height++
  if (height === 0) return null
  const kept = bottom - height
  let top = -1
  for (let col = 0; col < width; col++) {
    const own: number[] = []
    for (let row = 0; row <= kept; row++) if (records[row]![col]) own.push(row)
    if (col < left || col >= right) {
      if (own.length > 0) return null
      continue
    }
    if (own.length !== height || own.some((row, i) => row !== own[0]! + i)) return null
    if (top !== -1 && top !== own[0]) return null
    top = own[0]!
  }
  const edits: Edit[] = []
  for (let row = top; row <= kept; row++)
    for (let col = left; col < right; col++) {
      const from = grid[row + height]![col]!
      const to = grid[row]![col]!
      edits.push({ start: to.start, end: to.end, text: xml.slice(from.start, from.end) })
    }
  for (let row = kept + 1; row <= bottom; row++) edits.push(remove(rows[row]!))
  return { edits, applied: pending.length }
}

/** rejecting cells that came in with Shift cells down: one table at a time, each on the text the last left */
function undoDownShifts(
  xml: string,
  filter: RevisionXmlFilter | undefined,
): { xml: string; applied: number } {
  let out = xml
  let applied = 0
  let els = scan(out)
  for (let i = 0; i < els.length; i++) {
    if (els[i]!.name !== 'w:tbl') continue
    const shift = downShiftEdits(els, out, i, filter)
    if (!shift) continue
    const from = els[i]!.start
    out = applyEdits(out, shift.edits)
    applied += shift.applied
    els = scan(out)
    i = els.findIndex((el) => el.start === from)
  }
  return { xml: out, applied }
}

function applyEdits(xml: string, edits: Edit[]): string {
  const sorted = [...edits].sort((a, b) => a.start - b.start || b.end - a.end)
  // an edit inside a region another edit replaces is moot
  const kept: Edit[] = []
  let cutoff = -1
  for (const e of sorted) {
    if (e.start < cutoff) continue
    kept.push(e)
    cutoff = Math.max(cutoff, e.end)
  }
  let out = xml
  for (let i = kept.length - 1; i >= 0; i--) {
    const e = kept[i]!
    out = out.slice(0, e.start) + e.text + out.slice(e.end)
  }
  return out
}

/** pending tracked changes in a part (range markers not counted) */
export function countTrackedChangesXml(xml: string): number {
  if (
    !/<w:(?:ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge|numberingChange|\w+Change)[\s/>]/.test(
      xml,
    )
  )
    return 0
  const els = scan(xml)
  let n = 0
  for (const el of els) if (typeOf(el)) n++
  return n
}

export function applyTrackedChangesXml(
  xml: string,
  mode: RevisionXmlMode,
  filter?: RevisionXmlFilter,
): RevisionXmlResult {
  if (countTrackedChangesXml(xml) === 0) {
    // nothing countable is left, but a part whose move markup lost its partner
    // (the move was resolved earlier, e.g. in the body) still gets the orphan
    // markers back instead of shipping markup Word never resolves
    const cleaned = cleanMoveRangeMarkers(xml)
    return { xml: cleaned.xml, applied: 0, remaining: 0 }
  }
  const shifted =
    mode === 'reject' && xml.includes('<w:cellIns') ? undoDownShifts(xml, filter) : null
  const source = shifted?.xml ?? xml
  const els = scan(source)
  const { edits, applied } = planEdits(els, source, mode, filter)
  // a property holder emptied by independent edits (markers and record both resolved) goes
  const out = edits.length
    ? applyEdits(source, edits).replace(/<w:(rPr|trPr)><\/w:\1>/g, '')
    : source
  const cleaned = cleanMoveRangeMarkers(out)
  return {
    xml: cleaned.xml,
    applied: applied + (shifted?.applied ?? 0),
    remaining: countTrackedChangesXml(cleaned.xml),
  }
}

/** the editor's accept / reject decisions replayed on the part's original bytes */
export function applyRevisionResolutions(
  xml: string,
  resolutions: readonly RevisionResolution[] | undefined,
): string {
  let out = xml
  for (const r of resolutions ?? []) out = applyTrackedChangesXml(out, r.mode, r.filter).xml
  return out
}

export interface PropChangeRecord {
  author: string
  date?: string
  dateUtc?: string
  id?: string
  /** inner XML of the recorded previous properties (empty for records without one) */
  oldXml: string
  /** w:cellMerge only: the merge state after / before the revision */
  vMerge?: string
  vMergeOrig?: string
}

/**
 * The first `changeName` record (w:tblPrChange, w:trPrChange, w:tcPrChange, w:tblGridChange,
 * w:tblPrExChange, w:sectPrChange, w:numberingChange, w:cellMerge) in a property fragment.
 */
export function readPropChangeXml(xml: string, changeName: string): PropChangeRecord | undefined {
  if (!xml.includes(`<${changeName}`)) return undefined
  const els = scan(xml)
  const change = els.find((el) => el.name === changeName)
  if (!change) return undefined
  const holderName = PROP_CHANGE[changeName]
  const old = holderName ? children(els, change).find((c) => c.name === holderName) : undefined
  const date = attrOf(change, 'w:date')
  const dateUtc = attrOf(change, 'w16du:dateUtc')
  const id = attrOf(change, 'w:id')
  const vMerge = attrOf(change, 'w:vMerge')
  const vMergeOrig = attrOf(change, 'w:vMergeOrig')
  return {
    author: attrOf(change, 'w:author') ?? '',
    ...(date ? { date } : {}),
    ...(dateUtc ? { dateUtc } : {}),
    ...(id ? { id } : {}),
    oldXml: old ? xml.slice(old.innerStart, old.innerEnd) : '',
    ...(vMerge ? { vMerge } : {}),
    ...(vMergeOrig ? { vMergeOrig } : {}),
  }
}

/**
 * Resolve one property-change record inside a property fragment (a raw tblPr / trPr /
 * tcPr / tblGrid / tblPrEx / pPr-with-sectPr slice): accept drops the record, reject
 * puts the recorded previous properties back. Records without a previous-properties
 * payload (numberingChange, cellMerge) just drop.
 */
export function resolvePropChangeXml(
  xml: string,
  changeName: string,
  mode: RevisionXmlMode,
): string {
  if (!xml.includes(`<${changeName}`)) return xml
  const els = scan(xml)
  const change = els.find((el) => el.name === changeName)
  if (!change) return xml
  const edits =
    mode === 'accept' || !PROP_CHANGE[changeName]
      ? [remove(change)]
      : revertPropChange(els, xml, change)
  return applyEdits(xml, edits).replace(/<w:(rPr|trPr|tcPr|tblPrEx)><\/w:\1>/g, '')
}

/**
 * Drop orphaned move-range markers: an end whose start is gone (the moved
 * paragraph was accepted / rejected, so the wrapper left and the untouched
 * following paragraph keeps the end), and defensively a start whose end is
 * gone. A range pairs by side and w:id (a moveFrom range and a moveTo range
 * that happen to share an id are independent), and an end closes the nearest
 * still-open start of its side, mirroring how the parser closes the ranges it
 * reads. A range with both markers (a still-pending move) is left alone. The
 * final save runs this over the body and every story part because the editor
 * rebuilds edited paragraphs from its model, which does not model the range
 * markers.
 */
export function cleanMoveRangeMarkers(xml: string): { xml: string; removed: number } {
  if (!/<w:move(?:From|To)Range(?:Start|End)[\s/>]/.test(xml)) {
    return { xml, removed: 0 }
  }
  const els = scan(xml)
  const open = new Map<string, El[]>()
  const dropped: El[] = []
  for (const el of els) {
    if (!RANGE_MARKERS.has(el.name)) continue
    const side = el.name.startsWith('w:moveFrom') ? 'moveFrom' : 'moveTo'
    const idMatch = /\bw:id\s*=\s*(?:"([^"]*)"|'([^']*)')/.exec(el.attrs)
    const id = idMatch?.[1] ?? idMatch?.[2] ?? ''
    if (el.name.endsWith('Start')) {
      const key = side + '|' + id
      const stack = open.get(key)
      if (stack) stack.push(el)
      else open.set(key, [el])
    } else {
      // an end pairs with the nearest still-open start of the same side and
      // w:id; a shared id across moveFrom / moveTo is two independent ranges
      const start = open.get(side + '|' + id)?.pop()
      if (!start) dropped.push(el)
    }
  }
  for (const stack of open.values()) dropped.push(...stack)
  if (dropped.length === 0) return { xml, removed: 0 }
  return { xml: applyEdits(xml, dropped.map(remove)), removed: dropped.length }
}

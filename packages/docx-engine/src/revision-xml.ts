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
      case 'w:cellDel':
      case 'w:numberingChange':
        edits.push(remove(el))
        break
      default:
        edits.push(...(mode === 'accept' ? [remove(el)] : revertPropChange(els, xml, el)))
    }
  }
  return { edits, applied }
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

/** pending tracked changes in a part (range markers and cell merges not counted) */
export function countTrackedChangesXml(xml: string): number {
  if (!/<w:(?:ins|del|moveFrom|moveTo|cellIns|cellDel|numberingChange|\w+Change)[\s/>]/.test(xml))
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
  if (countTrackedChangesXml(xml) === 0) return { xml, applied: 0, remaining: 0 }
  const els = scan(xml)
  const { edits, applied } = planEdits(els, xml, mode, filter)
  const out = edits.length ? applyEdits(xml, edits) : xml
  return { xml: out, applied, remaining: countTrackedChangesXml(out) }
}

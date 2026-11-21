/**
 * The comparison of a story outside the body model (a header, a footer, a
 * footnote, an endnote, the cells of a table): the revised story's own XML
 * with the text differences written as w:ins / w:del.
 *
 * Only paragraphs made of plain text runs are rewritten. A paragraph that
 * holds a field, a picture, a hyperlink or a content control and differs is
 * left as the revised document has it and counted in `skipped`.
 */
import { escapeXmlAttr, escapeXmlText } from './xml-utils'

/** a piece of a compared text: kept, deleted (a range of the original) or inserted (a range of the revised) */
export interface StoryChunk {
  kind: 'keep' | 'del' | 'ins' | 'fmt'
  left?: [number, number]
  right?: [number, number]
  /** keep: the original text is tracked here, character for character over the revised range */
  carry?: boolean
}

export type StoryAlignment =
  | { kind: 'pair'; left: number; right: number }
  | { kind: 'removed'; left: number }
  | { kind: 'added'; right: number }

export interface StoryDiffContext {
  author: string
  date: string
  takeId: () => string
  /** how two lists of block texts line up */
  align: (left: string[], right: string[]) => StoryAlignment[]
  /** the differences of two paragraph texts; `tracked` says which characters sit in a tracked change */
  chunks: (
    left: string,
    right: string,
    tracked?: { left: ReadonlyArray<'' | 'del' | 'ins'>; right: ReadonlyArray<'' | 'del' | 'ins'> },
  ) => StoryChunk[]
  /** tracked changes either story already holds stay under their own author */
  keepRevisions?: boolean
  /** two texts that count as the same under the comparison's rules */
  same: (left: string, right: string) => boolean
  /** false leaves every table as the revised document has it */
  tables: boolean
  counts: { insertions: number; deletions: number; skipped: number; carried?: number }
}

interface Element {
  name: string
  xml: string
  /** the element's content, '' for a self-closing one */
  inner: string
  open: string
}

const TAG = /<(\/?)([A-Za-z_][\w:.-]*)((?:"[^"]*"|'[^']*'|[^>"'])*?)(\/?)>/g

/** the child elements of a fragment, text between them dropped */
export function childElements(xml: string): Element[] {
  const out: Element[] = []
  let depth = 0
  let start = 0
  let open = ''
  let innerStart = 0
  let name = ''
  TAG.lastIndex = 0
  for (let m = TAG.exec(xml); m; m = TAG.exec(xml)) {
    const closing = m[1] === '/'
    const selfClosing = m[4] === '/'
    if (closing) {
      depth--
      if (depth === 0)
        out.push({
          name,
          xml: xml.slice(start, m.index + m[0].length),
          inner: xml.slice(innerStart, m.index),
          open,
        })
      continue
    }
    if (depth === 0) {
      if (selfClosing) {
        out.push({ name: m[2]!, xml: m[0], inner: '', open: m[0] })
        continue
      }
      start = m.index
      open = m[0]
      innerStart = m.index + m[0].length
      name = m[2]!
    }
    if (!selfClosing) depth++
  }
  return out
}

const ATOM = '\uFFFC'

interface Piece {
  rPr: string
  text: string
  /** the element an ATOM stands for */
  raw?: string
  /** opening tags of the tracked change the text already sits in */
  ins?: string
  del?: string
}

interface SimpleParagraph {
  open: string
  pPr: string
  pieces: Piece[]
  bookmarkStarts: string
  bookmarkEnds: string
  text: string
}

const isTracked = (piece: Piece): boolean => piece.ins !== undefined || piece.del !== undefined

const trackedFlags = (p: SimpleParagraph): Array<'' | 'del' | 'ins'> =>
  p.pieces.flatMap((piece) =>
    new Array<'' | 'del' | 'ins'>(piece.text.length).fill(
      piece.ins !== undefined ? 'ins' : piece.del !== undefined ? 'del' : '',
    ),
  )

const unescapeXml = (text: string): string =>
  text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&amp;/g, '&')

/** run children that ride along as one opaque character */
const RUN_ATOMS = new Set([
  'w:footnoteRef',
  'w:endnoteRef',
  'w:separator',
  'w:continuationSeparator',
  'w:annotationRef',
  'w:noBreakHyphen',
  'w:softHyphen',
  'w:sym',
  'w:ptab',
])
const DROPPED = new Set(['w:proofErr', 'w:lastRenderedPageBreak'])
const RANGE_STARTS = new Set(['w:bookmarkStart', 'w:permStart'])
const RANGE_ENDS = new Set(['w:bookmarkEnd', 'w:permEnd'])
const MOVE_RANGES = new Set([
  'w:moveFromRangeStart',
  'w:moveFromRangeEnd',
  'w:moveToRangeStart',
  'w:moveToRangeEnd',
])

/**
 * The paragraph as text runs, or null when it holds more than text. Every
 * pending change reads as accepted, unless `markup` keeps them: deleted text
 * stays and each piece names the change it sits in.
 */
function simpleParagraph(p: Element, markup = false): SimpleParagraph | null {
  const out: SimpleParagraph = {
    open: p.inner === '' && p.xml.endsWith('/>') ? p.open.replace(/\/>$/, '>') : p.open,
    pPr: '',
    pieces: [],
    bookmarkStarts: '',
    bookmarkEnds: '',
    text: '',
  }
  const run = (r: Element, change: Pick<Piece, 'ins' | 'del'>): boolean => {
    let rPr = ''
    const push = (text: string, raw?: string) =>
      out.pieces.push({ rPr, text, ...(raw !== undefined ? { raw } : {}), ...change })
    for (const child of childElements(r.inner)) {
      if (child.name === 'w:rPr') rPr = child.xml
      else if (child.name === 'w:t' || child.name === 'w:delText') push(unescapeXml(child.inner))
      else if (child.name === 'w:tab') push('\t')
      else if (child.name === 'w:cr') push('\n')
      else if (child.name === 'w:br') {
        if (/\bw:type\s*=\s*["'](?:page|column)["']/.test(child.open)) push(ATOM, child.xml)
        else push('\n')
      } else if (RUN_ATOMS.has(child.name)) push(ATOM, child.xml)
      else if (!DROPPED.has(child.name)) return false
    }
    return true
  }
  // a move kept as markup reads as the deletion and the insertion it is made of
  const plainTag = (open: string) =>
    open.replace(/^<w:move(?:From|To)\b/, (m) => (m === '<w:moveTo' ? '<w:ins' : '<w:del'))
  const walk = (children: Element[], change: Pick<Piece, 'ins' | 'del'> = {}): boolean => {
    for (const child of children) {
      if (child.name === 'w:pPr') out.pPr = child.xml
      else if (child.name === 'w:r') {
        if (!run(child, change)) return false
      } else if (child.name === 'w:ins' || child.name === 'w:moveTo') {
        const inner = markup ? { ...change, ins: plainTag(child.open) } : change
        if (!walk(childElements(child.inner), inner)) return false
      } else if (child.name === 'w:del' || child.name === 'w:moveFrom') {
        if (!markup) continue
        if (!walk(childElements(child.inner), { ...change, del: plainTag(child.open) }))
          return false
      } else if (RANGE_STARTS.has(child.name)) out.bookmarkStarts += child.xml
      else if (RANGE_ENDS.has(child.name)) out.bookmarkEnds += child.xml
      else if (!DROPPED.has(child.name) && !MOVE_RANGES.has(child.name)) return false
    }
    return true
  }
  if (!walk(childElements(p.inner))) return null
  out.text = out.pieces.map((piece) => piece.text).join('')
  return out
}

/** what a paragraph that is more than text runs reads as, pending deletions left out unless `markup` keeps them */
function flatText(xml: string, markup = false): string {
  const live = markup
    ? xml
    : xml.replace(/<w:(del|moveFrom)\b(?:[^>/]|\/(?!>))*>[\s\S]*?<\/w:\1>/g, '')
  let out = ''
  const text = markup
    ? /<w:(?:t|delText)(?:\s[^>]*)?>([^<]*)<\/w:(?:t|delText)>|<w:tab\/>|<w:br\/>/g
    : /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:tab\/>|<w:br\/>/g
  for (const m of live.matchAll(text))
    out += m[1] !== undefined ? unescapeXml(m[1]) : m[0] === '<w:tab/>' ? '\t' : '\n'
  return out
}

/** every outermost w:p of a fragment, with where it sits */
function paragraphSpans(xml: string): Array<{ start: number; end: number }> {
  const out: Array<{ start: number; end: number }> = []
  let depth = 0
  let start = 0
  TAG.lastIndex = 0
  for (let m = TAG.exec(xml); m; m = TAG.exec(xml)) {
    if (m[2] !== 'w:p') continue
    if (m[1] === '/') {
      depth--
      if (depth === 0) out.push({ start, end: m.index + m[0].length })
    } else if (m[4] === '/') {
      if (depth === 0) out.push({ start: m.index, end: m.index + m[0].length })
    } else {
      if (depth === 0) start = m.index
      depth++
    }
  }
  return out
}

/** rows, cells and paragraphs in order: two tables of one shape pair paragraph by paragraph */
function tableShape(xml: string): string {
  let out = ''
  let inParagraph = 0
  TAG.lastIndex = 0
  for (let m = TAG.exec(xml); m; m = TAG.exec(xml)) {
    const name = m[2]!
    if (name === 'w:p') {
      if (m[1] === '/') inParagraph--
      else {
        if (inParagraph === 0) out += 'p'
        if (m[4] !== '/') inParagraph++
      }
    } else if (inParagraph === 0 && m[1] !== '/' && (name === 'w:tr' || name === 'w:tc')) {
      out += name === 'w:tr' ? '|r' : '|c'
    }
  }
  return out
}

const slicePieces = (pieces: Piece[], start: number, end: number): Piece[] => {
  const out: Piece[] = []
  let pos = 0
  for (const piece of pieces) {
    const from = Math.max(start, pos)
    const to = Math.min(end, pos + piece.text.length)
    if (to > from) out.push({ ...piece, text: piece.text.slice(from - pos, to - pos) })
    pos += piece.text.length
  }
  return out
}

function runXml(piece: Piece, deleted: boolean): string {
  if (piece.raw !== undefined) return deleted ? '' : `<w:r>${piece.rPr}${piece.raw}</w:r>`
  const tag = deleted ? 'w:delText' : 'w:t'
  let content = ''
  for (const part of piece.text.split(/([\t\n])/)) {
    if (part === '\t') content += '<w:tab/>'
    else if (part === '\n') content += '<w:br/>'
    else if (part) content += `<${tag} xml:space="preserve">${escapeXmlText(part)}</${tag}>`
  }
  return content ? `<w:r>${piece.rPr}${content}</w:r>` : ''
}

const revisionAttrs = (ctx: StoryDiffContext): string =>
  ` w:id="${escapeXmlAttr(ctx.takeId())}" w:author="${escapeXmlAttr(ctx.author)}" w:date="${escapeXmlAttr(ctx.date)}"`

function wrapped(kind: 'ins' | 'del', pieces: Piece[], ctx: StoryDiffContext): string {
  const runs = pieces.map((piece) => runXml(piece, kind === 'del')).join('')
  return runs ? `<w:${kind}${revisionAttrs(ctx)}>${runs}</w:${kind}>` : ''
}

/** a tracked change's opening tag under an id of this document */
const renumbered = (open: string, ctx: StoryDiffContext): string =>
  /\bw:id\s*=/.test(open)
    ? open.replace(/\bw:id\s*=\s*(?:"[^"]*"|'[^']*')/, `w:id="${escapeXmlAttr(ctx.takeId())}"`)
    : open.replace(/^<w:(ins|del)/, `<w:$1 w:id="${escapeXmlAttr(ctx.takeId())}"`)

/**
 * Writes pieces inside the changes they sit in: their own, or the comparison's
 * (`label`, one change for each call) where they have none. Neighbours in the
 * same change share one wrapper, and a change that opens a second time takes
 * a new id.
 */
function trackedWriter(ctx: StoryDiffContext) {
  const opened = new Set<string>()
  const once = (tag: string | undefined): string => {
    if (!tag) return ''
    if (!opened.has(tag)) {
      opened.add(tag)
      return tag
    }
    return renumbered(tag, ctx)
  }
  let out = ''
  let group: { ins?: string; del?: string; runs: string } | null = null
  const close = () => {
    if (!group) return
    out += `${once(group.ins)}${once(group.del)}${group.runs}${group.del ? '</w:del>' : ''}${group.ins ? '</w:ins>' : ''}`
    group = null
  }
  return {
    add(pieces: Piece[], label: 'ins' | 'del' | null): void {
      let labelTag: string | undefined
      const labelled = () => (labelTag ??= `<w:${label}${revisionAttrs(ctx)}>`)
      for (const piece of pieces) {
        const run = runXml(piece, piece.del !== undefined || label === 'del')
        if (!run) continue
        const ins = piece.ins ?? (label === 'ins' && !isTracked(piece) ? labelled() : undefined)
        const del = piece.del ?? (label === 'del' ? labelled() : undefined)
        if (group && group.ins === ins && group.del === del) group.runs += run
        else {
          close()
          group = { ...(ins ? { ins } : {}), ...(del ? { del } : {}), runs: run }
        }
      }
    },
    xml(): string {
      close()
      return out
    },
  }
}

/** the original's changes, renumbered once each however many pieces they span */
function carrier(ctx: StoryDiffContext): (piece: Piece) => Pick<Piece, 'ins' | 'del'> {
  const seen = new Map<string, string>()
  const one = (open: string): string => {
    let out = seen.get(open)
    if (!out) {
      out = renumbered(open, ctx)
      seen.set(open, out)
      ctx.counts.carried = (ctx.counts.carried ?? 0) + 1
    }
    return out
  }
  return (piece) => ({
    ...(piece.ins !== undefined ? { ins: one(piece.ins) } : {}),
    ...(piece.del !== undefined ? { del: one(piece.del) } : {}),
  })
}

/** the paragraph properties with the revision recorded on the paragraph mark */
function markedPPr(pPr: string, kind: 'ins' | 'del', ctx: StoryDiffContext): string {
  const record = `<w:${kind}${revisionAttrs(ctx)}/>`
  if (!pPr || /^<w:pPr(?:\s[^>]*)?\/>$/.test(pPr)) return `<w:pPr><w:rPr>${record}</w:rPr></w:pPr>`
  const live = pPr.indexOf('<w:pPrChange')
  const head = live === -1 ? pPr.length : live
  const emptyRPr = pPr.lastIndexOf('<w:rPr/>', head)
  if (emptyRPr !== -1)
    return `${pPr.slice(0, emptyRPr)}<w:rPr>${record}</w:rPr>${pPr.slice(emptyRPr + '<w:rPr/>'.length)}`
  const rPr = pPr.lastIndexOf('<w:rPr>', head)
  if (rPr !== -1) {
    const at = rPr + '<w:rPr>'.length
    return pPr.slice(0, at) + record + pPr.slice(at)
  }
  const sectPr = pPr.indexOf('<w:sectPr')
  const close = pPr.lastIndexOf('</w:pPr>')
  const at = [sectPr, live, close].filter((i) => i !== -1).sort((a, b) => a - b)[0]!
  return `${pPr.slice(0, at)}<w:rPr>${record}</w:rPr>${pPr.slice(at)}`
}

/** properties of a paragraph of the other package, without what points into that package */
const foreignPPr = (pPr: string): string =>
  pPr
    .replace(/<w:numPr>[\s\S]*?<\/w:numPr>/g, '')
    .replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, '')
    .replace(/<w:pPrChange\b[\s\S]*?<\/w:pPrChange>/g, '')

function changedParagraph(
  left: SimpleParagraph,
  right: SimpleParagraph,
  ctx: StoryDiffContext,
): string {
  let body = ''
  if (ctx.keepRevisions) {
    const carry = carrier(ctx)
    const writer = trackedWriter(ctx)
    const tracked = { left: trackedFlags(left), right: trackedFlags(right) }
    const fresh = (pieces: Piece[], deleted: boolean) =>
      pieces.some((piece) => !isTracked(piece) && runXml(piece, deleted))
    for (const chunk of ctx.chunks(left.text, right.text, tracked)) {
      if (chunk.kind === 'del') {
        const pieces = slicePieces(left.pieces, chunk.left![0], chunk.left![1])
        if (fresh(pieces, true)) ctx.counts.deletions++
        // text the original already tracks keeps that change instead of the comparison's
        const mine = pieces.map((piece) =>
          isTracked(piece) ? { ...piece, ...carry(piece) } : piece,
        )
        let i = 0
        while (i < mine.length) {
          let j = i
          while (j < mine.length && isTracked(mine[j]!) === isTracked(mine[i]!)) j++
          writer.add(mine.slice(i, j), isTracked(mine[i]!) ? null : 'del')
          i = j
        }
      } else if (chunk.kind === 'ins') {
        const pieces = slicePieces(right.pieces, chunk.right![0], chunk.right![1])
        if (fresh(pieces, false)) ctx.counts.insertions++
        writer.add(pieces, 'ins')
      } else if (chunk.carry) {
        let at = chunk.right![0]
        for (const mine of slicePieces(left.pieces, chunk.left![0], chunk.left![1])) {
          const to = at + mine.text.length
          writer.add(
            slicePieces(right.pieces, at, to).map((piece) =>
              isTracked(piece) || !isTracked(mine) ? piece : { ...piece, ...carry(mine) },
            ),
            null,
          )
          at = to
        }
      } else writer.add(slicePieces(right.pieces, chunk.right![0], chunk.right![1]), null)
    }
    return `${right.open}${right.pPr}${right.bookmarkStarts}${writer.xml()}${right.bookmarkEnds}</w:p>`
  }
  for (const chunk of ctx.chunks(left.text, right.text)) {
    if (chunk.kind === 'del') {
      const xml = wrapped('del', slicePieces(left.pieces, chunk.left![0], chunk.left![1]), ctx)
      if (xml) ctx.counts.deletions++
      body += xml
    } else if (chunk.kind === 'ins') {
      const xml = wrapped('ins', slicePieces(right.pieces, chunk.right![0], chunk.right![1]), ctx)
      if (xml) ctx.counts.insertions++
      body += xml
    } else {
      body += slicePieces(right.pieces, chunk.right![0], chunk.right![1])
        .map((piece) => runXml(piece, false))
        .join('')
    }
  }
  return `${right.open}${right.pPr}${right.bookmarkStarts}${body}${right.bookmarkEnds}</w:p>`
}

/** a paragraph that is one tracked change: every piece of it inserted or deleted */
const whollyTracked = (p: SimpleParagraph): boolean =>
  p.pieces.length > 0 ? p.pieces.every(isTracked) : liveRPr(p.pPr).search(PARA_MARK_CHANGE) !== -1

const PARA_MARK_CHANGE = /<w:(?:ins|del|moveFrom|moveTo)\b[^>]*\/>/g
const liveRPr = (pPr: string): string =>
  /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(pPr.split('<w:pPrChange')[0]!)?.[0] ?? ''

/** the paragraph-mark changes of a w:pPr of the other package under ids of this document */
export function renumberedParaMark(pPr: string, takeId: () => string): string {
  const rPr = liveRPr(pPr)
  if (!rPr) return pPr
  const next = rPr.replace(PARA_MARK_CHANGE, (tag) =>
    tag.replace(/\bw:id\s*=\s*(?:"[^"]*"|'[^']*')/, `w:id="${escapeXmlAttr(takeId())}"`),
  )
  return pPr.replace(rPr, next)
}

function wholeParagraph(
  p: SimpleParagraph,
  kind: 'ins' | 'del',
  ctx: StoryDiffContext,
  foreign: boolean,
): string {
  const open = foreign ? '<w:p>' : p.open
  const ranges = foreign ? ['', ''] : [p.bookmarkStarts, p.bookmarkEnds]
  if (ctx.keepRevisions) {
    const carry = carrier(ctx)
    const writer = trackedWriter(ctx)
    const pieces = foreign ? p.pieces.map((piece) => ({ ...piece, ...carry(piece) })) : p.pieces
    // a paragraph its own document inserted or deleted as a tracked change stays that change
    if (whollyTracked(p)) {
      writer.add(pieces, null)
      const pPr = foreign ? renumberedParaMark(foreignPPr(p.pPr), ctx.takeId) : p.pPr
      if (foreign && p.pieces.length === 0) ctx.counts.carried = (ctx.counts.carried ?? 0) + 1
      return `${open}${pPr}${ranges[0]}${writer.xml()}${ranges[1]}</w:p>`
    }
    if (kind === 'ins') ctx.counts.insertions++
    else ctx.counts.deletions++
    const pPr = markedPPr(foreign ? foreignPPr(p.pPr) : p.pPr, kind, ctx)
    writer.add(pieces, kind)
    return `${open}${pPr}${ranges[0]}${writer.xml()}${ranges[1]}</w:p>`
  }
  const pPr = markedPPr(foreign ? foreignPPr(p.pPr) : p.pPr, kind, ctx)
  if (kind === 'ins') ctx.counts.insertions++
  else ctx.counts.deletions++
  return `${open}${pPr}${ranges[0]}${wrapped(kind, p.pieces, ctx)}${ranges[1]}</w:p>`
}

/** one paragraph of each side that stand for each other */
function pairedParagraph(left: Element, right: Element, ctx: StoryDiffContext): string {
  const markup = ctx.keepRevisions === true
  const l = simpleParagraph(left, markup)
  const r = simpleParagraph(right, markup)
  const leftText = l ? l.text : flatText(left.xml, markup)
  const rightText = r ? r.text : flatText(right.xml, markup)
  // the same text still differs when the original tracks a change the revised one does not show
  const carries = markup && !!l && !!r && l.pieces.some(isTracked)
  if (ctx.same(leftText, rightText) && !carries) return right.xml
  if (!l || !r) {
    ctx.counts.skipped++
    return right.xml
  }
  return changedParagraph(l, r, ctx)
}

/**
 * Two tables that stand for each other: with the same rows, cells and
 * paragraphs, each cell paragraph is compared with its twin. null when the
 * shapes differ.
 */
export function diffTableXml(left: string, right: string, ctx: StoryDiffContext): string | null {
  if (tableShape(left) !== tableShape(right)) return null
  const mine = paragraphSpans(left)
  const theirs = paragraphSpans(right)
  if (mine.length !== theirs.length) return null
  let out = ''
  let at = 0
  theirs.forEach((span, i) => {
    const l = childElements(left.slice(mine[i]!.start, mine[i]!.end))[0]
    const r = childElements(right.slice(span.start, span.end))[0]
    out += right.slice(at, span.start)
    out += l && r ? pairedParagraph(l, r, ctx) : right.slice(span.start, span.end)
    at = span.end
  })
  return out + right.slice(at)
}

const TEXT_BOX = /(<w:txbxContent(?:\s[^>]*)?>)([\s\S]*?)(<\/w:txbxContent>)/g

/**
 * A block that draws text boxes, with the text of each box compared with the
 * box at the same place in the original. null when the two do not hold the
 * same number of boxes, or when no box differs.
 */
export function diffTextBoxXml(left: string, right: string, ctx: StoryDiffContext): string | null {
  const mine = [...left.matchAll(TEXT_BOX)]
  if (mine.length === 0 || mine.length !== [...right.matchAll(TEXT_BOX)].length) return null
  let at = 0
  const out = right.replace(TEXT_BOX, (_all, open: string, inner: string, close: string) => {
    return open + diffStoryXml(mine[at++]![2]!, inner, ctx) + close
  })
  return out === right ? null : out
}

const TABLE_TAG = '\u0001table '
const OPAQUE_TAG = '\u0001other '

function unitText(element: Element, markup: boolean): string {
  if (element.name === 'w:p')
    return simpleParagraph(element, markup)?.text ?? flatText(element.xml, markup)
  if (element.name === 'w:tbl') return TABLE_TAG + flatText(element.xml, markup)
  return OPAQUE_TAG + element.xml
}

/**
 * The revised story (the content of a w:hdr, w:ftr, w:footnote or w:endnote)
 * with what differs from the original marked as tracked changes.
 */
export function diffStoryXml(left: string, right: string, ctx: StoryDiffContext): string {
  const mine = childElements(left)
  const theirs = childElements(right)
  const markup = ctx.keepRevisions === true
  const leftTexts = mine.map((element) => unitText(element, markup))
  const rightTexts = theirs.map((element) => unitText(element, markup))
  let out = ''
  for (const entry of ctx.align(leftTexts, rightTexts)) {
    if (entry.kind === 'removed') {
      const element = mine[entry.left]!
      const p = element.name === 'w:p' ? simpleParagraph(element, markup) : null
      if (p) out += wholeParagraph(p, 'del', ctx, true)
      else if (element.name === 'w:p' || element.name === 'w:tbl') ctx.counts.skipped++
      continue
    }
    const element = theirs[entry.right]!
    if (entry.kind === 'added') {
      const p = element.name === 'w:p' ? simpleParagraph(element, markup) : null
      if (p) out += wholeParagraph(p, 'ins', ctx, false)
      else {
        if (element.name === 'w:p' || element.name === 'w:tbl') ctx.counts.skipped++
        out += element.xml
      }
      continue
    }
    const twin = mine[entry.left]!
    if (element.name === 'w:p' && twin.name === 'w:p') out += pairedParagraph(twin, element, ctx)
    else if (element.name === 'w:tbl' && twin.name === 'w:tbl') {
      if (!ctx.tables || ctx.same(leftTexts[entry.left]!, rightTexts[entry.right]!))
        out += element.xml
      else {
        const diffed = diffTableXml(twin.xml, element.xml, ctx)
        if (diffed === null) ctx.counts.skipped++
        out += diffed ?? element.xml
      }
    } else {
      if (leftTexts[entry.left] !== rightTexts[entry.right]) ctx.counts.skipped++
      out += element.xml
    }
  }
  return out
}

/** the whole of a story the original does not have, marked inserted where it is plain text */
export function insertedStoryXml(right: string, ctx: StoryDiffContext): string {
  return childElements(right)
    .map((element) => {
      const p = element.name === 'w:p' ? simpleParagraph(element, ctx.keepRevisions) : null
      return p && p.pieces.some((piece) => piece.raw === undefined && piece.text.trim())
        ? wholeParagraph(p, 'ins', ctx, false)
        : element.xml
    })
    .join('')
}

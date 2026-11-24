/**
 * The comparison of a story outside the body model (a header, a footer, a
 * footnote, an endnote, the cells of a table): the revised story's own XML
 * with the text differences written as w:ins / w:del, paragraph formatting
 * as w:pPrChange, and added or removed table rows as tracked rows.
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

/** a stretch the comparison deleted or inserted as one change of its own */
export interface StoryFragment {
  kind: 'ins' | 'del'
  /** w:id of the w:ins / w:del around the runs */
  id: string
  /** w:id of the record on the paragraph mark, when the whole paragraph is the change */
  markId?: string
  text: string
}

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
  counts: {
    insertions: number
    deletions: number
    skipped: number
    carried?: number
    formatChanges?: number
  }
  /**
   * The revised paragraph's w:pPr with the formatting change the comparison
   * records on it, null when the two paragraphs are formatted alike. Absent,
   * paragraph formatting is not compared.
   */
  paragraphFormat?: (
    left: string | undefined,
    right: string | undefined,
  ) => { xml: string; carried?: boolean } | null
  /** collects what the comparison deletes and inserts, for the caller to read moves from */
  fragments?: StoryFragment[]
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

const revisionAttrs = (ctx: StoryDiffContext, id = ctx.takeId()): string =>
  ` w:id="${escapeXmlAttr(id)}" w:author="${escapeXmlAttr(ctx.author)}" w:date="${escapeXmlAttr(ctx.date)}"`

function wrapped(
  kind: 'ins' | 'del',
  pieces: Piece[],
  ctx: StoryDiffContext,
  markId?: string,
): string {
  const runs = pieces.map((piece) => runXml(piece, kind === 'del')).join('')
  if (!runs) return ''
  const id = ctx.takeId()
  if (ctx.fragments && pieces.every((piece) => piece.raw === undefined))
    ctx.fragments.push({
      kind,
      id,
      ...(markId ? { markId } : {}),
      text: pieces.map((piece) => piece.text).join(''),
    })
  return `<w:${kind}${revisionAttrs(ctx, id)}>${runs}</w:${kind}>`
}

/**
 * A deleted and an inserted stretch of one text relabelled as a move: the
 * wrappers become w:moveFrom / w:moveTo inside a range of that name, a
 * paragraph mark that went or came with them likewise. Only the labels
 * change, so accepting or rejecting resolves as before.
 */
export function relabelledMove(
  xml: string,
  fragment: StoryFragment,
  name: string,
  ctx: Pick<StoryDiffContext, 'author' | 'date' | 'takeId'>,
): string {
  const side = fragment.kind === 'del' ? 'moveFrom' : 'moveTo'
  const open = new RegExp(`<w:${fragment.kind}( w:id="${fragment.id}"[^>]*)>`)
  const at = open.exec(xml)
  if (!at || at[1]!.endsWith('/')) return xml
  const close = xml.indexOf(`</w:${fragment.kind}>`, at.index)
  if (close === -1) return xml
  const range = escapeXmlAttr(ctx.takeId())
  let out =
    xml.slice(0, at.index) +
    `<w:${side}RangeStart w:id="${range}" w:author="${escapeXmlAttr(ctx.author)}" w:date="${escapeXmlAttr(ctx.date)}" w:name="${escapeXmlAttr(name)}"/>` +
    `<w:${side}${at[1]}>` +
    // moved-away text stays w:t, as Word writes it
    xml.slice(at.index + at[0].length, close).replace(/(<\/?)w:delText\b/g, '$1w:t') +
    `</w:${side}><w:${side}RangeEnd w:id="${range}"/>` +
    xml.slice(close + `</w:${fragment.kind}>`.length)
  if (fragment.markId)
    out = out.replace(
      new RegExp(`<w:${fragment.kind}( w:id="${fragment.markId}"[^>]*/>)`),
      `<w:${side}$1`,
    )
  return out
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
function markedPPr(pPr: string, kind: 'ins' | 'del', ctx: StoryDiffContext, id?: string): string {
  const record = `<w:${kind}${revisionAttrs(ctx, id)}/>`
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
  const markId = ctx.takeId()
  const pPr = markedPPr(foreign ? foreignPPr(p.pPr) : p.pPr, kind, ctx, markId)
  if (kind === 'ins') ctx.counts.insertions++
  else ctx.counts.deletions++
  return `${open}${pPr}${ranges[0]}${wrapped(kind, p.pieces, ctx, markId)}${ranges[1]}</w:p>`
}

const ownPPr = (p: Element): string | undefined => {
  const first = /^\s*<w:pPr[\s/>]/.test(p.inner) ? childElements(p.inner)[0] : undefined
  return first?.name === 'w:pPr' ? first.xml : undefined
}

/** the revised paragraph's w:pPr once the comparison recorded a formatting change on it; null when it records none */
function comparedPPr(left: Element, right: Element, ctx: StoryDiffContext): string | null {
  const change = ctx.paragraphFormat?.(ownPPr(left), ownPPr(right))
  if (!change) return null
  if (change.carried) ctx.counts.carried = (ctx.counts.carried ?? 0) + 1
  else ctx.counts.formatChanges = (ctx.counts.formatChanges ?? 0) + 1
  return change.xml
}

/** the paragraph's own XML under another w:pPr */
function withPPr(p: Element, pPr: string): string {
  const open = p.inner === '' && p.xml.endsWith('/>') ? p.open.replace(/\/>$/, '>') : p.open
  const own = ownPPr(p)
  const at = own ? p.inner.indexOf(own) : 0
  return `${open}${p.inner.slice(0, at)}${pPr}${p.inner.slice(at + (own?.length ?? 0))}</w:p>`
}

/** one paragraph of each side that stand for each other */
function pairedParagraph(left: Element, right: Element, ctx: StoryDiffContext): string {
  if (left.xml === right.xml) return right.xml
  const markup = ctx.keepRevisions === true
  const l = simpleParagraph(left, markup)
  const r = simpleParagraph(right, markup)
  const leftText = l ? l.text : flatText(left.xml, markup)
  const rightText = r ? r.text : flatText(right.xml, markup)
  const pPr = comparedPPr(left, right, ctx)
  // the same text still differs when the original tracks a change the revised one does not show
  const carries = markup && !!l && !!r && l.pieces.some(isTracked)
  if (ctx.same(leftText, rightText) && !carries)
    return pPr === null ? right.xml : withPPr(right, pPr)
  if (!l || !r) {
    ctx.counts.skipped++
    return pPr === null ? right.xml : withPPr(right, pPr)
  }
  return changedParagraph(l, pPr === null ? r : { ...r, pPr }, ctx)
}

/** drop the drawings, VML pictures, embedded objects and references of a fragment from the other package */
export function withoutForeignObjects(xml: string): string {
  return xml
    .replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, '')
    .replace(/<w:pict(?:\s[^>]*)?>[\s\S]*?<\/w:pict>/g, '')
    .replace(/<w:object(?:\s[^>]*)?>[\s\S]*?<\/w:object>/g, '')
    .replace(/<w:(?:footnote|endnote|comment)Reference\b[^>]*\/>/g, '')
    .replace(/<w:commentRange(?:Start|End)\b[^>]*\/>/g, '')
}

const REVISION_TAG =
  /<w:(?:ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge|numberingChange|\w+Change)\b[^>]*>/g

/** a fragment of the other package that comes along with the changes it tracks, under ids of this document */
function carriedXml(xml: string, ctx: StoryDiffContext): string {
  return withoutForeignObjects(xml)
    .replace(/<w:numPr>[\s\S]*?<\/w:numPr>/g, '')
    .replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, '')
    .replace(REVISION_TAG, (tag) =>
      tag.replace(/\bw:id\s*=\s*(?:"[^"]*"|'[^']*')/, `w:id="${escapeXmlAttr(ctx.takeId())}"`),
    )
}

/** each paragraph of a table fragment diffed against the one at the same place; the shapes are equal */
function diffParagraphsInPlace(left: string, right: string, ctx: StoryDiffContext): string {
  const mine = paragraphSpans(left)
  const theirs = paragraphSpans(right)
  let out = ''
  let at = 0
  theirs.forEach((span, i) => {
    const l = mine[i] ? childElements(left.slice(mine[i].start, mine[i].end))[0] : undefined
    const r = childElements(right.slice(span.start, span.end))[0]
    out += right.slice(at, span.start)
    out += l && r ? pairedParagraph(l, r, ctx) : right.slice(span.start, span.end)
    at = span.end
  })
  return out + right.slice(at)
}

/** every paragraph of a fragment rewritten, what lies between them kept */
function mapParagraphs(xml: string, rewrite: (p: Element) => string): string {
  let out = ''
  let at = 0
  for (const span of paragraphSpans(xml)) {
    const p = childElements(xml.slice(span.start, span.end))[0]
    out += xml.slice(at, span.start) + (p ? rewrite(p) : xml.slice(span.start, span.end))
    at = span.end
  }
  return out + xml.slice(at)
}

interface TableRow {
  el: Element
  /** the row's own tracked insertion or deletion (w:trPr) */
  change: 'ins' | 'del' | null
  text: string
}

const ROW_CHANGE = /<w:(ins|del)\b[^>]*\/>/
const CELL_CHANGE = /<w:(?:cellIns|cellDel|cellMerge)\b[^>]*\/>/

const ownProps = (el: Element, name: string): Element | undefined =>
  childElements(el.inner).find((child) => child.name === name)

/** the part of a properties element that is in force, its change history left out */
const liveProps = (props: Element | undefined): string =>
  props ? props.inner.split(/<w:t[rc]PrChange\b/)[0]! : ''

const cellsOf = (row: Element): Element[] =>
  childElements(row.inner).filter((child) => child.name === 'w:tc')

const cellText = (cell: Element, markup: boolean): string =>
  paragraphSpans(cell.xml)
    .map((span) => {
      const p = childElements(cell.xml.slice(span.start, span.end))[0]!
      return simpleParagraph(p, markup)?.text ?? flatText(p.xml, markup)
    })
    .join('\n')

function tableRows(
  xml: string,
  markup: boolean,
): { table: Element; children: Element[]; rows: TableRow[] } | null {
  const table = childElements(xml)[0]
  if (!table || table.name !== 'w:tbl') return null
  const rows: TableRow[] = []
  const children = childElements(table.inner)
  for (const child of children) {
    if (child.name === 'w:sdt' || child.name === 'w:customXml') return null
    if (child.name !== 'w:tr') continue
    const change = ROW_CHANGE.exec(liveProps(ownProps(child, 'w:trPr')))?.[1] as
      'ins' | 'del' | undefined
    const text = cellsOf(child)
      .map((cell) => cellText(cell, markup))
      .join('\t')
    rows.push({ el: child, change: change ?? null, text })
  }
  return { table, children, rows }
}

const rebuilt = (el: Element, children: string[]): string =>
  `${el.open}${children.join('')}</${el.name}>`

/** an element's properties child with `record` added; one of the same kind already there gives way */
function withRecord(el: Element, propsName: string, record: string, existing: RegExp): string {
  const children = childElements(el.inner)
  const at = children.findIndex((child) => child.name === propsName)
  if (at === -1) {
    const first = children.findIndex((child) => child.name !== 'w:tblPrEx')
    const xml = children.map((child) => child.xml)
    xml.splice(first === -1 ? xml.length : first, 0, `<${propsName}>${record}</${propsName}>`)
    return rebuilt(el, xml)
  }
  const props = children[at]!
  const live = liveProps(props).replace(existing, '')
  const history = props.inner.slice(liveProps(props).length)
  return rebuilt(
    el,
    children.map((child, i) =>
      i === at ? `<${propsName}>${live}${record}${history}</${propsName}>` : child.xml,
    ),
  )
}

/** counts a row as one change however many paragraphs it holds */
const quiet = (ctx: StoryDiffContext): StoryDiffContext => ({
  ...ctx,
  counts: { insertions: 0, deletions: 0, skipped: 0 },
  fragments: undefined,
})

/** a paragraph that is more than text runs, as the text it reads as */
const flatParagraph = (p: Element, markup: boolean): SimpleParagraph => ({
  open: '<w:p>',
  pPr: '',
  pieces: [{ rPr: '', text: flatText(p.xml, markup) }],
  bookmarkStarts: '',
  bookmarkEnds: '',
  text: flatText(p.xml, markup),
})

/** a row only the original has: every paragraph deleted and the row marked deleted */
function goneRow(row: TableRow, ctx: StoryDiffContext): string {
  const markup = ctx.keepRevisions === true
  if (markup && row.change) {
    ctx.counts.carried = (ctx.counts.carried ?? 0) + 1
    return carriedXml(row.el.xml, ctx)
  }
  ctx.counts.deletions++
  const silent = quiet(ctx)
  const xml = mapParagraphs(withoutForeignObjects(row.el.xml), (p) =>
    wholeParagraph(simpleParagraph(p, markup) ?? flatParagraph(p, markup), 'del', silent, true),
  )
  return withRecord(childElements(xml)[0]!, 'w:trPr', `<w:del${revisionAttrs(ctx)}/>`, ROW_CHANGE)
}

/** a row only the revised document has: every paragraph inserted and the row marked inserted */
function addedRow(row: TableRow, ctx: StoryDiffContext): string {
  const markup = ctx.keepRevisions === true
  if (markup && row.change) return row.el.xml
  ctx.counts.insertions++
  const silent = quiet(ctx)
  const xml = mapParagraphs(row.el.xml, (p) => {
    const simple = simpleParagraph(p, markup)
    return simple ? wholeParagraph(simple, 'ins', silent, false) : p.xml
  })
  return withRecord(childElements(xml)[0]!, 'w:trPr', `<w:ins${revisionAttrs(ctx)}/>`, ROW_CHANGE)
}

const TRACKING = /^w:(?:ins|del|cellIns|cellDel|cellMerge|\w+Change)$/

/**
 * `el` with the property change the original tracks on its `propsName`
 * child (`changeName`), when the revised element has the same properties in
 * force and tracks no change of its own there: rejected, it goes back to
 * what the original recorded. Unchanged otherwise.
 */
function withCarriedPropChange(
  left: Element,
  el: Element,
  propsName: string,
  changeName: string,
  ctx: StoryDiffContext,
): Element {
  const mine = ownProps(left, propsName)
  const record = mine && childElements(mine.inner).find((child) => child.name === changeName)
  if (!record) return el
  const theirs = ownProps(el, propsName)
  const inForce = (props: Element | undefined) =>
    childElements(props?.inner ?? '')
      .filter((child) => !TRACKING.test(child.name))
      .map((child) => child.xml)
      .join('')
  const own = childElements(theirs?.inner ?? '').some((child) => child.name === changeName)
  if (own || inForce(mine) !== inForce(theirs)) return el
  ctx.counts.carried = (ctx.counts.carried ?? 0) + 1
  const carried = carriedXml(record.xml, ctx)
  if (!theirs) return childElements(withRecord(el, propsName, carried, /$^/))[0]!
  const at = el.inner.indexOf(theirs.xml)
  const props = `${theirs.open.replace(/\/>$/, '>')}${theirs.inner}${carried}</${propsName}>`
  return childElements(
    `${el.open}${el.inner.slice(0, at)}${props}${el.inner.slice(at + theirs.xml.length)}</${el.name}>`,
  )[0]!
}

/** the row and cell changes the original tracks and the revised row does not show */
function carriedRowRecords(left: Element, rightXml: string, ctx: StoryDiffContext): string {
  let row = childElements(rightXml)[0]!
  const carried = (tag: string) => {
    ctx.counts.carried = (ctx.counts.carried ?? 0) + 1
    return carriedXml(tag, ctx)
  }
  const mine = cellsOf(left)
  const theirs = cellsOf(row)
  if (mine.length === theirs.length) {
    let at = 0
    const cells = childElements(row.inner).map((child) => {
      if (child.name !== 'w:tc') return child.xml
      const twin = mine[at++]!
      const cell = withCarriedPropChange(twin, child, 'w:tcPr', 'w:tcPrChange', ctx)
      const record = CELL_CHANGE.exec(liveProps(ownProps(twin, 'w:tcPr')))?.[0]
      if (!record || CELL_CHANGE.test(liveProps(ownProps(cell, 'w:tcPr')))) return cell.xml
      return withRecord(cell, 'w:tcPr', carried(record), CELL_CHANGE)
    })
    row = childElements(rebuilt(row, cells))[0]!
  }
  row = withCarriedPropChange(left, row, 'w:trPr', 'w:trPrChange', ctx)
  const record = ROW_CHANGE.exec(liveProps(ownProps(left, 'w:trPr')))?.[0]
  if (record && !ROW_CHANGE.test(liveProps(ownProps(row, 'w:trPr'))))
    return withRecord(row, 'w:trPr', carried(record), ROW_CHANGE)
  return row.xml
}

/** a cell of each side that stand for each other: the revised cell, its content compared */
function pairedCell(left: Element, right: Element, ctx: StoryDiffContext): string {
  const content = (cell: Element) => {
    const props = ownProps(cell, 'w:tcPr')
    return props ? cell.inner.replace(props.xml, '') : cell.inner
  }
  const body = diffStoryXml(content(left), content(right), ctx)
  return `${right.open}${ownProps(right, 'w:tcPr')?.xml ?? ''}${body}</w:tc>`
}

/** a cell by the grid columns it covers */
interface GridCell {
  start: number
  span: number
  /** its part in a vertical merge */
  merge: 'restart' | 'continue' | null
}

/** a cell by the entries of the column alignment it covers, its first one and what kind they all are */
interface PlacedCell {
  index: number
  from: number
  span: number
  kind: StoryAlignment['kind']
  merge: GridCell['merge']
}

/**
 * How the columns of two tables stand for each other when one was added or
 * removed. Each side is a grid every row fills (cells merged across or down
 * count by the columns they cover), and no merged cell reaches across a
 * column only the other side has. null otherwise, or when no column of one
 * stands for a column of the other.
 */
interface Columns {
  aligned: StoryAlignment[]
  /** the cells of each row, by the entries of `aligned` they stand at */
  placed: Map<Element, PlacedCell[]>
  /** `<w:tcW>` of a column, for a cell a row has to be given */
  widths: { left: string[]; right: string[] }
  /** a cell merged down in a table whose rows are not all paired */
  tall: boolean
}

const UNEVEN = /<w:(?:hMerge|gridBefore|gridAfter)\b/

function gridCells(row: Element): GridCell[] | null {
  if (UNEVEN.test(liveProps(ownProps(row, 'w:trPr')))) return null
  const cells: GridCell[] = []
  let start = 0
  for (const cell of cellsOf(row)) {
    const props = liveProps(ownProps(cell, 'w:tcPr'))
    if (UNEVEN.test(props)) return null
    const span = Number(/<w:gridSpan\b[^>]*\bw:val="(\d+)"/.exec(props)?.[1] ?? 1) || 1
    const merge = /<w:vMerge\b([^>]*)>/.exec(props)
    cells.push({
      start,
      span,
      merge: !merge ? null : /w:val="restart"/.test(merge[1]!) ? 'restart' : 'continue',
    })
    start += span
  }
  return cells
}

const gridCols = (table: Element): Element[] | null => {
  const grid = ownProps(table, 'w:tblGrid')
  return grid ? childElements(grid.inner).filter((child) => child.name === 'w:gridCol') : null
}

/** every cell that continues a vertical merge has the cell it continues right above it */
const chained = (rows: GridCell[][]): boolean =>
  rows.every((row, r) =>
    row.every(
      (cell) =>
        cell.merge !== 'continue' ||
        !!rows[r - 1]?.some((up) => up.start === cell.start && up.span === cell.span && up.merge),
    ),
  )

function alignedColumns(
  mine: { table: Element; rows: TableRow[] },
  theirs: { table: Element; rows: TableRow[] },
  pairs: Array<{ left: number; right: number }>,
  whole: boolean,
  ctx: StoryDiffContext,
  /** the paired rows hold as many cells on both sides: columns of equal number stand for each other in order */
  inOrder = false,
): Columns | null {
  const grid = (rows: TableRow[]) => {
    const cells = rows.map((row) => gridCells(row.el))
    const widthOf = (row: GridCell[] | null) => row?.reduce((sum, cell) => sum + cell.span, 0)
    const width = widthOf(cells[0] ?? null)
    return width && cells.every((row) => widthOf(row) === width)
      ? { width, cells: cells as GridCell[][] }
      : null
  }
  const left = grid(mine.rows)
  const right = grid(theirs.rows)
  if (!left || !right) return null
  // a cell merged down takes part in every row it covers
  const merged = (cells: GridCell[][]) => cells.some((row) => row.some((cell) => cell.merge))
  const tall = !whole && (merged(left.cells) || merged(right.cells))
  if (tall && !(chained(left.cells) && chained(right.cells))) return null
  const markup = ctx.keepRevisions === true
  // a column reads as the cells that start in it
  const texts = (side: 'left' | 'right') => {
    const { rows, cells, width } =
      side === 'left' ? { rows: mine.rows, ...left } : { rows: theirs.rows, ...right }
    return Array.from({ length: width }, (_, col) =>
      pairs
        .map((pair) => {
          const at = cells[pair[side]]!.findIndex((cell) => cell.start === col)
          return at === -1 ? '' : cellText(cellsOf(rows[pair[side]]!.el)[at]!, markup)
        })
        .join('\n'),
    )
  }
  const aligned: StoryAlignment[] =
    inOrder && left.width === right.width
      ? Array.from({ length: left.width }, (_, i) => ({ kind: 'pair', left: i, right: i }))
      : ctx.align(texts('left'), texts('right'))
  if (!aligned.some((entry) => entry.kind === 'pair')) return null

  const entryOf = { left: new Array<number>(left.width), right: new Array<number>(right.width) }
  aligned.forEach((entry, i) => {
    if (entry.kind !== 'added') entryOf.left[entry.left] = i
    if (entry.kind !== 'removed') entryOf.right[entry.right] = i
  })
  /** a row's cells by the entries they cover; null when one reaches across a column of the other side */
  const place = (cells: GridCell[], at: number[]): PlacedCell[] | null => {
    const out: PlacedCell[] = []
    for (const [index, cell] of cells.entries()) {
      const from = at[cell.start]!
      const kind = aligned[from]!.kind
      for (let k = 1; k < cell.span; k++) {
        if (at[cell.start + k] !== from + k || aligned[from + k]!.kind !== kind) return null
      }
      out.push({ index, from, span: cell.span, kind, merge: cell.merge })
    }
    return out
  }
  const placed = new Map<Element, PlacedCell[]>()
  for (const [side, rows, cells] of [
    ['left', mine.rows, left.cells],
    ['right', theirs.rows, right.cells],
  ] as const) {
    for (const [i, row] of rows.entries()) {
      const cellsAt = place(cells[i]!, entryOf[side])
      if (!cellsAt) return null
      placed.set(row.el, cellsAt)
    }
  }
  // the cells two paired rows share are merged alike
  for (const pair of pairs) {
    const shared = (row: TableRow) =>
      placed
        .get(row.el)!
        .filter((cell) => cell.kind === 'pair')
        .map((cell) => `${cell.from}:${cell.span}:${cell.merge}`)
        .join(' ')
    if (shared(mine.rows[pair.left]!) !== shared(theirs.rows[pair.right]!)) return null
  }

  const widths = (table: Element, rows: TableRow[], cells: GridCell[][], width: number) => {
    const cols = gridCols(table)
    return Array.from({ length: width }, (_, col) => {
      for (const [i, row] of cells.entries()) {
        const at = row.findIndex((cell) => cell.start === col && cell.span === 1)
        const own =
          at === -1
            ? null
            : /<w:tcW\b[^>]*\/>/.exec(liveProps(ownProps(cellsOf(rows[i]!.el)[at]!, 'w:tcPr')))
        if (own) return own[0]
      }
      const w = cols?.length === width ? /\bw:w="(\d+)"/.exec(cols[col]!.open)?.[1] : null
      return w ? `<w:tcW w:w="${w}" w:type="dxa"/>` : ''
    })
  }
  return {
    aligned,
    placed,
    tall,
    widths: {
      left: widths(mine.table, mine.rows, left.cells, left.width),
      right: widths(theirs.table, theirs.rows, right.cells, right.width),
    },
  }
}

/** a cell of a column only the original has: its text deleted, the cell marked deleted */
function goneCell(cell: Element, ctx: StoryDiffContext): string {
  const markup = ctx.keepRevisions === true
  const silent = quiet(ctx)
  const xml = mapParagraphs(withoutForeignObjects(cell.xml), (p) =>
    wholeParagraph(simpleParagraph(p, markup) ?? flatParagraph(p, markup), 'del', silent, true),
  )
  return withRecord(
    childElements(xml)[0]!,
    'w:tcPr',
    `<w:cellDel${revisionAttrs(ctx)}/>`,
    CELL_CHANGE,
  )
}

/** a cell of a column only the revised table has: its text inserted, the cell marked inserted */
function addedCell(cell: Element, ctx: StoryDiffContext): string {
  const markup = ctx.keepRevisions === true
  const silent = quiet(ctx)
  const xml = mapParagraphs(cell.xml, (p) => {
    const simple = simpleParagraph(p, markup)
    return simple ? wholeParagraph(simple, 'ins', silent, false) : p.xml
  })
  return withRecord(
    childElements(xml)[0]!,
    'w:tcPr',
    `<w:cellIns${revisionAttrs(ctx)}/>`,
    CELL_CHANGE,
  )
}

/** a cell an unpaired row is given inside a vertical merge of the other table: it continues the merge */
interface MergedStandIn {
  span: number
  width: string
}

/** an empty cell standing in a row for a column the row's own table does not have */
const standInCell = (
  width: string,
  kind: 'cellIns' | 'cellDel',
  ctx: StoryDiffContext,
  merged?: MergedStandIn,
): string =>
  `<w:tc><w:tcPr>${merged?.width ?? width}${merged && merged.span > 1 ? `<w:gridSpan w:val="${merged.span}"/>` : ''}${merged ? '<w:vMerge/>' : ''}<w:${kind}${revisionAttrs(ctx)}/></w:tcPr><w:p/></w:tc>`

/** the rows of the compared table in the order they are written; a paired row is both */
interface WrittenRow {
  left?: Element
  right?: Element
}

/**
 * With rows only one table has among cells merged down: the cells such a row
 * is given where a merge of the other table runs through it. null when a
 * merge both tables have would be cut by a row of one of them.
 */
function mergedStandIns(
  order: WrittenRow[],
  columns: Columns,
): Map<Element, Map<number, MergedStandIn>> | null {
  const out = new Map<Element, Map<number, MergedStandIn>>()
  for (const [k, row] of order.entries()) {
    const above = order[k - 1]
    for (const el of [row.left, row.right]) {
      if (!el) continue
      for (const cell of columns.placed.get(el)!) {
        // a row of one table that starts a merge the rows below continue: not followed
        if (cell.merge === 'restart' && cell.kind === 'pair' && !(row.left && row.right))
          return null
        if (cell.merge !== 'continue' || cell.kind !== 'pair') continue
        const up = above && columns.placed.get((above.left ?? above.right)!)
        if (!up?.some((c) => c.from === cell.from && c.span === cell.span && c.merge)) return null
      }
    }
    if (row.left && row.right) continue
    const other = row.left ? 'right' : 'left'
    const below = order.slice(k + 1).find((next) => next[other])?.[other]
    if (!below) continue
    const cells = cellsOf(below)
    const runs = new Map<number, MergedStandIn>()
    for (const cell of columns.placed.get(below)!) {
      if (cell.merge !== 'continue' || cell.kind === 'pair') continue
      const own = /<w:tcW\b[^>]*\/>/.exec(liveProps(ownProps(cells[cell.index]!, 'w:tcPr')))
      runs.set(cell.from, { span: cell.span, width: own?.[0] ?? '' })
    }
    if (runs.size > 0) out.set((row.left ?? row.right)!, runs)
  }
  return out
}

/** a row with its cells replaced, what is not a cell kept before or after them as it stood */
function withCells(row: Element, cells: string[]): string {
  const children = childElements(row.inner)
  const first = children.findIndex((child) => child.name === 'w:tc')
  const other = (from: number, to: number) =>
    children
      .slice(from, to)
      .filter((child) => child.name !== 'w:tc')
      .map((child) => child.xml)
  return rebuilt(row, [...other(0, first), ...cells, ...other(first, children.length)])
}

/**
 * A row only one table has, given a cell for each column the other table
 * alone has, so the column is whole: accepted or rejected, it goes or stays
 * in every row alike.
 */
function withStandIns(
  xml: string,
  source: Element,
  side: 'left' | 'right',
  columns: Columns,
  ctx: StoryDiffContext,
  merged?: Map<number, MergedStandIn>,
): string {
  const row = childElements(xml)[0]!
  const own = cellsOf(row)
  const starts = new Map(columns.placed.get(source)!.map((cell) => [cell.from, cell]))
  const cells: string[] = []
  let covered = 0
  columns.aligned.forEach((entry, i) => {
    if (i < covered) return
    const cell = starts.get(i)
    const run = merged?.get(i)
    if (run) covered = i + run.span
    if (cell) {
      // among cells merged down a column goes or stays by its cells alone: the row's own cell is marked too
      const kind = entry.kind === 'added' ? 'cellIns' : entry.kind === 'removed' ? 'cellDel' : null
      const mine = own[cell.index]!
      cells.push(
        columns.tall && kind && !CELL_CHANGE.test(liveProps(ownProps(mine, 'w:tcPr')))
          ? withRecord(mine, 'w:tcPr', `<w:${kind}${revisionAttrs(ctx)}/>`, CELL_CHANGE)
          : mine.xml,
      )
    } else if (entry.kind === 'added' && side === 'left')
      cells.push(standInCell(columns.widths.right[entry.right]!, 'cellIns', ctx, run))
    else if (entry.kind === 'removed' && side === 'right')
      cells.push(standInCell(columns.widths.left[entry.left]!, 'cellDel', ctx, run))
  })
  return withCells(row, cells)
}

/** the revised grid with the columns only the original has put back where they stood */
function mergedGrid(left: Element, right: Element, columns: Columns): string | null {
  const grid = ownProps(right, 'w:tblGrid')
  const mine = gridCols(left)
  const theirs = gridCols(right)
  if (!grid || !mine || !theirs) return null
  if (mine.length !== columns.widths.left.length || theirs.length !== columns.widths.right.length)
    return null
  const rest = childElements(grid.inner).filter((child) => child.name !== 'w:gridCol')
  return rebuilt(grid, [
    ...columns.aligned.map((entry) =>
      entry.kind === 'removed' ? mine[entry.left]!.xml : theirs[entry.right]!.xml,
    ),
    ...rest.map((child) => child.xml),
  ])
}

/** a row of each side that stand for each other */
function pairedRow(
  left: TableRow,
  right: TableRow,
  ctx: StoryDiffContext,
  columns: Columns | null,
): string {
  // with columns aligned, a row that reads the same still owes a cell to each column one side lacks
  if (!columns && left.el.xml === right.el.xml) return right.el.xml
  let xml: string
  const mine = cellsOf(left.el)
  if (columns) {
    const theirs = cellsOf(right.el)
    const from = (row: Element) =>
      new Map(columns.placed.get(row)!.map((cell) => [cell.from, cell]))
    const leftAt = from(left.el)
    const rightAt = from(right.el)
    const cells: string[] = []
    columns.aligned.forEach((entry, i) => {
      const old = leftAt.get(i)
      const now = rightAt.get(i)
      if (entry.kind === 'pair') {
        if (old && now) cells.push(pairedCell(mine[old.index]!, theirs[now.index]!, ctx))
      } else if (entry.kind === 'removed') {
        if (old) cells.push(goneCell(mine[old.index]!, ctx))
      } else if (now) cells.push(addedCell(theirs[now.index]!, ctx))
    })
    xml = withCells(right.el, cells)
  } else if (tableShape(left.el.xml) === tableShape(right.el.xml)) {
    xml = diffParagraphsInPlace(left.el.xml, right.el.xml, ctx)
  } else {
    let at = 0
    xml = rebuilt(
      right.el,
      childElements(right.el.inner).map((child) =>
        child.name === 'w:tc' ? pairedCell(mine[at++]!, child, ctx) : child.xml,
      ),
    )
  }
  return ctx.keepRevisions ? carriedRowRecords(left.el, xml, ctx) : xml
}

/**
 * Two tables that stand for each other. With the same rows, cells and
 * paragraphs each cell paragraph is compared with its twin. Otherwise the
 * rows are aligned by their text: a row only one side has is a tracked row
 * deletion or insertion, and a row both have is compared cell by cell. With
 * a column added or removed the columns are aligned the same way, and a
 * column only one side has is its cells tracked as inserted or deleted; a
 * merged cell that reaches across such a column leaves the table to the caller. null when no row of one stands for a row of the other.
 */
export function diffTableXml(left: string, right: string, ctx: StoryDiffContext): string | null {
  const xml = diffTableRows(left, right, ctx)
  if (xml === null || !ctx.keepRevisions || !/<w:tbl(?:Pr|Grid)Change\b/.test(left)) return xml
  // the original's tracked table properties and grid, on a table that still has them
  const mine = childElements(left)[0]
  let table = childElements(xml)[0]
  if (!mine || !table || table.name !== 'w:tbl') return xml
  table = withCarriedPropChange(mine, table, 'w:tblPr', 'w:tblPrChange', ctx)
  if (ownProps(table, 'w:tblGrid'))
    table = withCarriedPropChange(mine, table, 'w:tblGrid', 'w:tblGridChange', ctx)
  return table.xml
}

function diffTableRows(left: string, right: string, ctx: StoryDiffContext): string | null {
  const markup = ctx.keepRevisions === true
  const mine = tableRows(left, markup)
  const theirs = tableRows(right, markup)
  if (!mine || !theirs)
    return tableShape(left) === tableShape(right) ? diffParagraphsInPlace(left, right, ctx) : null
  // with every change accepted, a row pending deletion is not part of the table
  const struck = (rows: TableRow[]) => !markup && rows.some((row) => row.change === 'del')
  const inPlace =
    tableShape(left) === tableShape(right) && !struck(mine.rows) && !struck(theirs.rows)
  const live = (rows: TableRow[]) =>
    inPlace ? rows : rows.filter((row) => markup || row.change !== 'del')
  const leftRows = live(mine.rows)
  const rightRows = live(theirs.rows)
  const aligned: StoryAlignment[] = inPlace
    ? leftRows.map((_, i) => ({ kind: 'pair', left: i, right: i }))
    : ctx.align(
        leftRows.map((row) => row.text),
        rightRows.map((row) => row.text),
      )
  if (!aligned.some((entry) => entry.kind === 'pair')) return null
  // a column added or removed: the cells of a row stand for each other by their column
  const pairs = aligned.filter((entry) => entry.kind === 'pair')
  const sameCells = pairs.every(
    (entry) =>
      cellsOf(leftRows[entry.left]!.el).length === cellsOf(rightRows[entry.right]!.el).length,
  )
  const whole =
    pairs.length === aligned.length &&
    leftRows.length === mine.rows.length &&
    rightRows.length === theirs.rows.length
  // rows only one table has among cells merged down: the merges are followed through them
  const mergedDown = (rows: TableRow[]) =>
    rows.some((row) =>
      cellsOf(row.el).some((cell) => /<w:vMerge\b/.test(liveProps(ownProps(cell, 'w:tcPr')))),
    )
  const tall = !whole && (mergedDown(leftRows) || mergedDown(rightRows))
  const columns =
    sameCells && !tall
      ? null
      : alignedColumns(
          { table: mine.table, rows: leftRows },
          { table: theirs.table, rows: rightRows },
          pairs,
          whole,
          ctx,
          sameCells,
        )
  if (!columns && (tall || !sameCells)) return null
  let standIns: Map<Element, Map<number, MergedStandIn>> | null = null
  if (columns?.tall) {
    const order: WrittenRow[] = []
    let gone: WrittenRow[] = []
    for (const entry of aligned) {
      if (entry.kind === 'removed') gone.push({ left: leftRows[entry.left]!.el })
      else {
        const right = rightRows[entry.right]!.el
        order.push(
          ...gone,
          entry.kind === 'pair' ? { left: leftRows[entry.left]!.el, right } : { right },
        )
        gone = []
      }
    }
    standIns = mergedStandIns([...order, ...gone], columns)
    if (!standIns) return null
  }
  for (const entry of columns?.aligned ?? []) {
    if (entry.kind === 'added') ctx.counts.insertions++
    else if (entry.kind === 'removed') ctx.counts.deletions++
  }

  const rewritten = new Map<Element, string>()
  // a row the revised table strikes is gone once its changes are accepted
  for (const row of theirs.rows) if (!rightRows.includes(row)) rewritten.set(row.el, '')
  let pending = ''
  let last: Element | null = null
  for (const entry of aligned) {
    if (entry.kind === 'removed') {
      const row = leftRows[entry.left]!
      const gone = goneRow(row, ctx)
      pending += columns
        ? withStandIns(gone, row.el, 'left', columns, ctx, standIns?.get(row.el))
        : gone
      continue
    }
    const row = rightRows[entry.right]!
    let xml: string
    if (entry.kind === 'pair') xml = pairedRow(leftRows[entry.left]!, row, ctx, columns)
    else {
      const added = addedRow(row, ctx)
      xml = columns
        ? withStandIns(added, row.el, 'right', columns, ctx, standIns?.get(row.el))
        : added
    }
    if (pending || xml !== row.el.xml) rewritten.set(row.el, pending + xml)
    pending = ''
    last = row.el
  }
  if (pending && last) rewritten.set(last, (rewritten.get(last) ?? last.xml) + pending)
  if (rewritten.size === 0) return right
  const grid = columns ? mergedGrid(mine.table, theirs.table, columns) : null
  return rebuilt(
    theirs.table,
    theirs.children.map((child) =>
      grid && child.name === 'w:tblGrid' ? grid : (rewritten.get(child) ?? child.xml),
    ),
  )
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
      const sameText = ctx.same(leftTexts[entry.left]!, rightTexts[entry.right]!)
      if (!ctx.tables || (sameText && !ctx.paragraphFormat)) out += element.xml
      else {
        const diffed = diffTableXml(twin.xml, element.xml, ctx)
        if (diffed === null && !sameText) ctx.counts.skipped++
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

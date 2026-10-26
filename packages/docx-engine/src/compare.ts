/**
 * Compare two documents the way Word's Review > Compare does: the revised
 * document's content with the differences marked as tracked insertions and
 * deletions, so the result can be accepted or rejected like any other redline.
 *
 * The revision marks are the ones the rest of the suite already writes. A run
 * carries `ins` / `del` (see Run in types.ts) and the generator wraps
 * consecutive runs of the same revision in w:ins / w:del (see generate.ts);
 * a paragraph inserted or deleted as a whole is wrapped in a block-level
 * w:ins / w:del, the shape `docs apply --track` produces. A redline written
 * here therefore resolves through the same accept / reject path.
 *
 * Out of scope: w:rPrChange / w:pPrChange records and moved text. A run whose
 * formatting changed is reported as a deletion of the old run plus an
 * insertion of the new one, which is honest but more verbose than Word.
 */
import { saveDocx, type ParsedDocFull, type SaveBlock } from './patch'
import type { Block, GeneratedBlock, RevisionInfo, Run } from './types'

/** at what granularity a changed paragraph is diffed */
export type CompareGranularity = 'word' | 'char'

/** what the comparison leaves out */
export type CompareIgnore = 'formatting' | 'case' | 'whitespace'

export interface CompareOptions {
  /** author recorded on every revision; default "GenOffice" */
  author?: string
  /** ISO timestamp recorded on every revision; default: now */
  date?: string
  /** default "word" */
  granularity?: CompareGranularity
  /** default: nothing is ignored */
  ignore?: readonly CompareIgnore[]
}

export interface CompareSummary {
  /** paragraphs the revised document adds */
  insertedParagraphs: number
  /** paragraphs the revised document drops */
  deletedParagraphs: number
  /** paragraphs both have, with a different content */
  changedParagraphs: number
  /** paragraphs both have, identical */
  unchangedParagraphs: number
  /** inserted text regions (w:ins) */
  insertions: number
  /** deleted text regions (w:del) */
  deletions: number
}

export interface CompareResult {
  /** the redline document */
  bytes: Uint8Array
  summary: CompareSummary
}

interface Rules {
  ignoreCase: boolean
  ignoreWhitespace: boolean
  ignoreFormatting: boolean
}

/** one comparable unit of a document: a text paragraph, or a table / picture block */
interface Unit {
  /** index into parsed.blocks */
  index: number
  block: Block
  /** the comparable text: the paragraph text, or a table's cell texts */
  text: string
  /** a text paragraph, diffed word by word */
  textual: boolean
}

type Aligned =
  | { kind: 'pair'; left: Unit; right: Unit }
  | { kind: 'removed'; left: Unit }
  | { kind: 'added'; right: Unit }

/** a piece of the compared paragraph: text kept, deleted or inserted */
interface Chunk {
  kind: 'keep' | 'del' | 'ins'
  /** char range in the original text */
  left?: [number, number]
  /** char range in the revised text */
  right?: [number, number]
}

interface Token {
  text: string
  /** char offset in the paragraph text */
  start: number
  end: number
}

/** below this share of common words two paragraphs are unrelated, not an edit */
const PAIR_THRESHOLD = 0.6

/** the author recorded when the caller names none */
export const DEFAULT_COMPARE_AUTHOR = 'GenOffice'

/**
 * Build the redline: the revised document with the differences marked as
 * tracked insertions and deletions.
 */
export async function compareDocuments(
  original: ParsedDocFull,
  revised: ParsedDocFull,
  options: CompareOptions = {},
): Promise<CompareResult> {
  const plan = compareBlocks(original, revised, options)
  const bytes = await saveDocx(revised, plan.blocks, { revisionIdFloor: plan.nextRevisionId })
  return { bytes, summary: plan.summary }
}

/** the save plan of a comparison, for a caller that writes the file itself */
export function compareBlocks(
  original: ParsedDocFull,
  revised: ParsedDocFull,
  options: CompareOptions = {},
): { blocks: SaveBlock[]; summary: CompareSummary; nextRevisionId: number } {
  const rules: Rules = {
    ignoreCase: options.ignore?.includes('case') ?? false,
    ignoreWhitespace: options.ignore?.includes('whitespace') ?? false,
    ignoreFormatting: options.ignore?.includes('formatting') ?? false,
  }
  const author = options.author?.trim() || DEFAULT_COMPARE_AUTHOR
  const date = options.date || new Date().toISOString()
  const granularity = options.granularity === 'char' ? 'char' : 'word'
  const ids = new RevisionIds(revised.internal.documentXml)

  const left = unitsOf(original)
  const right = unitsOf(revised)
  const aligned = alignUnits(left, right, rules)

  const blocks: SaveBlock[] = []
  const summary: CompareSummary = {
    insertedParagraphs: 0,
    deletedParagraphs: 0,
    changedParagraphs: 0,
    unchangedParagraphs: 0,
    insertions: 0,
    deletions: 0,
  }

  for (const entry of aligned) {
    if (entry.kind === 'pair') {
      const leftBlock = entry.left.block
      const rightBlock = entry.right.block
      if (leftBlock.runs && rightBlock.runs) {
        const leftRuns = comparableRuns(leftBlock.runs)
        const rightRuns = comparableRuns(rightBlock.runs)
        const chunks = paragraphChunks(leftRuns, rightRuns, granularity, rules)
        if (!hasRevision(chunks)) {
          summary.unchangedParagraphs++
          blocks.push(originalPassthrough(rightBlock))
          continue
        }
        summary.changedParagraphs++
        const runs: Run[] = []
        for (const chunk of chunks) {
          if (chunk.kind === 'keep') {
            runs.push(...sliceRuns(rightRuns, chunk.right![0], chunk.right![1]).map(emittedRun))
            continue
          }
          const info: RevisionInfo = { author, date, id: ids.take() }
          if (chunk.kind === 'del') {
            summary.deletions++
            runs.push(
              ...sliceRuns(leftRuns, chunk.left![0], chunk.left![1]).map((r) => ({
                ...foreignRun(r),
                del: info,
              })),
            )
          } else {
            summary.insertions++
            runs.push(
              ...sliceRuns(rightRuns, chunk.right![0], chunk.right![1]).map((r) => ({
                ...emittedRun(r),
                ins: info,
              })),
            )
          }
        }
        blocks.push({ kind: 'generated', block: generatedFrom(rightBlock, runs) })
        continue
      }
      // a table, a picture or a field: the whole block is kept or replaced
      if (comparableText(entry.left.text, rules) === comparableText(entry.right.text, rules)) {
        summary.unchangedParagraphs++
        blocks.push(originalPassthrough(rightBlock))
        continue
      }
      summary.deletedParagraphs++
      summary.insertedParagraphs++
      blocks.push(deletedBlock(leftBlock, author, date, ids))
      blocks.push(insertedBlock(rightBlock, author, date, ids))
      continue
    }
    if (entry.kind === 'removed') {
      summary.deletedParagraphs++
      summary.deletions++
      blocks.push(deletedBlock(entry.left.block, author, date, ids))
      continue
    }
    summary.insertedParagraphs++
    summary.insertions++
    blocks.push(insertedBlock(entry.right.block, author, date, ids))
  }

  return { blocks, summary, nextRevisionId: ids.peek() }
}

/** the comparable units of a document, in body order */
function unitsOf(parsed: ParsedDocFull): Unit[] {
  const out: Unit[] = []
  parsed.blocks.forEach((block, index) => {
    if (block.hidden || block.blockRevision?.kind === 'del') return
    const textual = Array.isArray(block.runs)
    out.push({ index, block, textual, text: contentSignature(block) })
  })
  return out
}

/** the text a block is compared by: its runs, a table's cell texts, or its preview */
function contentSignature(block: Block): string {
  if (block.runs)
    return comparableRuns(block.runs)
      .map((r) => r.text)
      .join('')
  if (block.table) {
    return block.table.rows
      .map((row) => row.map((cell) => cell.paras.join('\n')).join('\t'))
      .join('\n')
  }
  // a picture is its bytes: the drawing xml names a media part by an id the
  // other package does not share, so the data url is what two pictures are
  // compared by. A block with nothing to compare by never matches another.
  if (block.imageDataUrl) return `${PICTURE_TOKEN}${block.imageDataUrl}`
  if (block.previewText) return block.previewText
  return `${PICTURE_TOKEN}opaque:${block.originalXml ?? ''}`
}

/**
 * Align the two documents: unchanged paragraphs anchor the sequence, and the
 * paragraphs between two anchors are paired by how much they have in common,
 * so an edited paragraph reads as an edit and a paragraph that really was
 * deleted stays a deletion.
 */
function alignUnits(left: Unit[], right: Unit[], rules: Rules): Aligned[] {
  const anchors = anchorPairs(
    left.map((u) => comparableText(u.text, rules)),
    right.map((u) => comparableText(u.text, rules)),
  )
  const out: Aligned[] = []
  let li = 0
  let ri = 0
  for (const [l, r] of anchors) {
    out.push(...pairRegion(left.slice(li, l), right.slice(ri, r)))
    out.push({ kind: 'pair', left: left[l]!, right: right[r]! })
    li = l + 1
    ri = r + 1
  }
  out.push(...pairRegion(left.slice(li), right.slice(ri)))
  return out
}

/** cells of the LCS table a gap between two unique anchors may cost before it is left unanchored */
const LCS_GAP_CELLS = 4_000_000

/**
 * The unchanged paragraphs that anchor the alignment: paragraphs whose text
 * occurs once in each document pair up first (their order is the longest
 * increasing subsequence), and the gaps between them are aligned by LCS. That
 * keeps the LCS table to the gaps instead of the whole document.
 */
function anchorPairs(left: string[], right: string[]): Array<[number, number]> {
  const count = (texts: string[]): Map<string, number> => {
    const m = new Map<string, number>()
    for (const t of texts) m.set(t, (m.get(t) ?? 0) + 1)
    return m
  }
  const leftCount = count(left)
  const rightCount = count(right)
  const rightIndex = new Map<string, number>()
  right.forEach((t, j) => {
    if (rightCount.get(t) === 1 && leftCount.get(t) === 1) rightIndex.set(t, j)
  })
  const candidates: Array<[number, number]> = []
  left.forEach((t, i) => {
    const j = rightIndex.get(t)
    if (j !== undefined) candidates.push([i, j])
  })
  const unique = longestIncreasing(candidates)

  const out: Array<[number, number]> = []
  let li = 0
  let ri = 0
  const fillGap = (lEnd: number, rEnd: number): void => {
    const n = lEnd - li
    const m = rEnd - ri
    if (n > 0 && m > 0 && n * m <= LCS_GAP_CELLS) {
      for (const [l, r] of lcsPairs(left.slice(li, lEnd), right.slice(ri, rEnd), (a, b) => a === b))
        out.push([li + l, ri + r])
    }
  }
  for (const [l, r] of unique) {
    fillGap(l, r)
    out.push([l, r])
    li = l + 1
    ri = r + 1
  }
  fillGap(left.length, right.length)
  return out
}

/** the longest run of candidate pairs whose right index rises with the left one (pairs come sorted by left) */
function longestIncreasing(pairs: Array<[number, number]>): Array<[number, number]> {
  const tails: number[] = []
  const tailAt: number[] = []
  const prev: number[] = new Array<number>(pairs.length).fill(-1)
  pairs.forEach(([, r], k) => {
    let lo = 0
    let hi = tails.length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (tails[mid]! < r) lo = mid + 1
      else hi = mid
    }
    tails[lo] = r
    tailAt[lo] = k
    prev[k] = lo > 0 ? tailAt[lo - 1]! : -1
  })
  const out: Array<[number, number]> = []
  for (let k = tailAt[tails.length - 1] ?? -1; k !== -1; k = prev[k]!) out.push(pairs[k]!)
  return out.reverse()
}

/** the text two paragraphs are matched by, the ignored differences set aside */
function comparableText(text: string, rules: Rules): string {
  let x = text
  if (rules.ignoreWhitespace) x = x.replace(/\s+/g, ' ').trim()
  if (rules.ignoreCase) x = x.toLowerCase()
  return x
}

/** the paragraphs between two anchors: pair the ones that resemble each other */
function pairRegion(left: Unit[], right: Unit[]): Aligned[] {
  if (!left.length) return right.map((u) => ({ kind: 'added' as const, right: u }))
  if (!right.length) return left.map((u) => ({ kind: 'removed' as const, left: u }))
  const n = left.length
  const m = right.length
  if (n * m > 250_000) {
    return [
      ...left.map((u) => ({ kind: 'removed' as const, left: u })),
      ...right.map((u) => ({ kind: 'added' as const, right: u })),
    ]
  }
  const score = (a: Unit, b: Unit): number => similarity(a.text, b.text)
  const dp: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0))
  const paired: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(-1))
  for (let i = 1; i <= n; i++) {
    for (let j = 1; j <= m; j++) {
      let best = dp[i - 1]![j]!
      let from = -1
      if (dp[i]![j - 1]! > best) {
        best = dp[i]![j - 1]!
        from = -2
      }
      const sim = score(left[i - 1]!, right[j - 1]!)
      if (sim >= PAIR_THRESHOLD && dp[i - 1]![j - 1]! + sim > best) {
        best = dp[i - 1]![j - 1]! + sim
        from = -3
      }
      dp[i]![j] = best
      paired[i]![j] = from
    }
  }
  const out: Aligned[] = []
  let i = n
  let j = m
  const region: Aligned[] = []
  while (i > 0 || j > 0) {
    const from = i > 0 && j > 0 ? paired[i]![j]! : -1
    if (from === -3) {
      region.push({ kind: 'pair', left: left[i - 1]!, right: right[j - 1]! })
      i--
      j--
    } else if (from === -2) {
      region.push({ kind: 'added', right: right[j - 1]! })
      j--
    } else if (i > 0) {
      region.push({ kind: 'removed', left: left[i - 1]! })
      i--
    } else {
      region.push({ kind: 'added', right: right[j - 1]! })
      j--
    }
  }
  region.reverse()
  out.push(...region)
  return out
}

/** the share of words two texts have in common, 0 to 1 */
function similarity(a: string, b: string): number {
  const left = tokenize(a, 'word').map((t) => t.text)
  const right = tokenize(b, 'word').map((t) => t.text)
  if (!left.length && !right.length) return 1
  if (!left.length || !right.length) return 0
  const common = lcsPairs(left, right, (x, y) => x === y).length
  return (2 * common) / (left.length + right.length)
}

/** the pieces of a changed paragraph: what stays, what was deleted, what was added */
function paragraphChunks(
  leftRuns: Run[],
  rightRuns: Run[],
  granularity: CompareGranularity,
  rules: Rules,
): Chunk[] {
  const leftText = leftRuns.map((r) => r.text).join('')
  const rightText = rightRuns.map((r) => r.text).join('')
  const leftTokens = tokenize(leftText, granularity)
  const rightTokens = tokenize(rightText, granularity)
  const pairs = lcsPairs(leftTokens, rightTokens, (a, b) => tokenEqual(a.text, b.text, rules))

  type Op =
    { kind: 'eq'; l: Token; r: Token } | { kind: 'del'; l: Token } | { kind: 'ins'; r: Token }
  const ops: Op[] = []
  let li = 0
  let ri = 0
  for (const [l, r] of pairs) {
    while (li < l) ops.push({ kind: 'del', l: leftTokens[li++]! })
    while (ri < r) ops.push({ kind: 'ins', r: rightTokens[ri++]! })
    ops.push({ kind: 'eq', l: leftTokens[l]!, r: rightTokens[r]! })
    li = l + 1
    ri = r + 1
  }
  while (li < leftTokens.length) ops.push({ kind: 'del', l: leftTokens[li++]! })
  while (ri < rightTokens.length) ops.push({ kind: 'ins', r: rightTokens[ri++]! })

  const leftStarts = runStarts(leftRuns)
  const rightStarts = runStarts(rightRuns)
  const chunks: Chunk[] = []
  let keepLeft: [number, number] | null = null
  let keepRight: [number, number] | null = null
  const flush = (): void => {
    if (keepLeft && keepRight) chunks.push({ kind: 'keep', left: keepLeft, right: keepRight })
    keepLeft = null
    keepRight = null
  }
  for (const op of ops) {
    if (op.kind === 'del') {
      // a whitespace run the caller asked to ignore is not a difference
      if (rules.ignoreWhitespace && !op.l.text.trim()) continue
      flush()
      chunks.push({ kind: 'del', left: [op.l.start, op.l.end] })
      continue
    }
    if (op.kind === 'ins') {
      if (rules.ignoreWhitespace && !op.r.text.trim()) continue
      flush()
      chunks.push({ kind: 'ins', right: [op.r.start, op.r.end] })
      continue
    }
    const styled =
      !rules.ignoreFormatting &&
      styleOf(leftRuns, leftStarts, op.l) !== styleOf(rightRuns, rightStarts, op.r)
    if (styled) {
      flush()
      chunks.push({ kind: 'del', left: [op.l.start, op.l.end] })
      chunks.push({ kind: 'ins', right: [op.r.start, op.r.end] })
      continue
    }
    if (keepLeft && keepRight) {
      keepLeft = [keepLeft[0], op.l.end]
      keepRight = [keepRight[0], op.r.end]
    } else {
      keepLeft = [op.l.start, op.l.end]
      keepRight = [op.r.start, op.r.end]
    }
  }
  flush()
  return mergeChunks(chunks)
}

/** neighbouring deletions (or insertions) become one revision region */
function mergeChunks(chunks: Chunk[]): Chunk[] {
  const out: Chunk[] = []
  for (const chunk of chunks) {
    const last = out[out.length - 1]
    if (last && last.kind === chunk.kind && chunk.kind !== 'keep') {
      const lastRange = chunk.kind === 'del' ? last.left! : last.right!
      const range = chunk.kind === 'del' ? chunk.left! : chunk.right!
      if (lastRange[1] === range[0]) {
        lastRange[1] = range[1]
        continue
      }
    }
    out.push({ ...chunk })
  }
  return out
}

function hasRevision(chunks: Chunk[]): boolean {
  return chunks.some((c) => c.kind !== 'keep')
}

/** words (or characters) and the whitespace between them */
function tokenize(text: string, granularity: CompareGranularity): Token[] {
  const re = granularity === 'char' ? /[^\s]|\s+/g : /\S+|\s+/g
  const out: Token[] = []
  for (const m of text.matchAll(re)) {
    out.push({ text: m[0], start: m.index, end: m.index + m[0].length })
  }
  return out
}

/** the longest common subsequence of two sequences, as index pairs */
function lcsPairs<T>(
  a: readonly T[],
  b: readonly T[],
  eq: (x: T, y: T) => boolean,
): Array<[number, number]> {
  const n = a.length
  const m = b.length
  const w = m + 1
  const lcs = new Int32Array((n + 1) * w)
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      lcs[i * w + j] = eq(a[i]!, b[j]!)
        ? lcs[(i + 1) * w + j + 1]! + 1
        : Math.max(lcs[(i + 1) * w + j]!, lcs[i * w + j + 1]!)
    }
  }
  const out: Array<[number, number]> = []
  let i = 0
  let j = 0
  while (i < n && j < m) {
    if (eq(a[i]!, b[j]!)) {
      out.push([i, j])
      i++
      j++
    } else if (lcs[(i + 1) * w + j]! >= lcs[i * w + j + 1]!) i++
    else j++
  }
  return out
}

/** two tokens are the same when the ignored differences are set aside */
function tokenEqual(a: string, b: string, rules: Rules): boolean {
  let x = a
  let y = b
  if (rules.ignoreWhitespace) {
    x = x.replace(/\s+/g, ' ')
    y = y.replace(/\s+/g, ' ')
  }
  if (rules.ignoreCase) {
    x = x.toLowerCase()
    y = y.toLowerCase()
  }
  return x === y
}

/** the runs covering a char range of the paragraph text */
function sliceRuns(runs: Run[], start: number, end: number): Run[] {
  const out: Run[] = []
  let pos = 0
  for (const run of runs) {
    const runStart = pos
    const runEnd = pos + run.text.length
    pos = runEnd
    if (runEnd <= start || runStart >= end) continue
    const from = Math.max(start, runStart)
    const to = Math.min(end, runEnd)
    const text = run.text.slice(from - runStart, to - runStart)
    if (!text) continue
    out.push({ ...run, text })
  }
  return out
}

/** the char offset each run starts at */
function runStarts(runs: Run[]): number[] {
  const out: number[] = []
  let pos = 0
  for (const run of runs) {
    out.push(pos)
    pos += run.text.length
  }
  return out
}

/** the formatting of the run a token starts in */
function styleOf(runs: Run[], starts: number[], token: Token): string {
  const index = runIndexAt(starts, token.start)
  const run = index >= 0 ? runs[index] : undefined
  if (!run) return ''
  return JSON.stringify([
    run.styleId ?? null,
    run.bold ?? false,
    run.italic ?? false,
    run.underline ?? false,
    run.strike ?? false,
    run.dstrike ?? false,
    run.color ?? null,
    run.sizeHalfPoints ?? null,
    run.font ?? null,
    run.fontAscii ?? null,
    run.eastAsiaFont ?? null,
    run.highlight ?? null,
    run.shading ?? null,
    run.vertAlign ?? null,
    run.caps ?? null,
    run.vanish ?? false,
    run.charSpacingTwips ?? null,
    run.charScalePct ?? null,
    run.rtl ?? false,
  ])
}

/** the run that holds a char offset */
function runIndexAt(starts: number[], offset: number): number {
  let lo = 0
  let hi = starts.length - 1
  let ans = -1
  while (lo <= hi) {
    const mid = (lo + hi) >> 1
    if (starts[mid]! <= offset) {
      ans = mid
      lo = mid + 1
    } else hi = mid - 1
  }
  return ans
}

/** the block as the generator wants it, with the given runs */
function generatedFrom(block: Block, runs: Run[]): GeneratedBlock {
  const type: GeneratedBlock['type'] =
    block.type === 'heading' ? 'heading' : block.type === 'listItem' ? 'listItem' : 'paragraph'
  return {
    type,
    ...(block.level !== undefined ? { level: block.level } : {}),
    ...(block.outlineOnly ? { outlineOnly: true } : {}),
    ...(block.styleId !== undefined ? { styleId: block.styleId } : {}),
    ...(block.list ? { list: block.list } : {}),
    ...(block.format ? { format: block.format } : {}),
    ...(block.rawPPr !== undefined ? { rawPPr: block.rawPPr } : {}),
    ...(block.bookmarks ? { bookmarks: block.bookmarks } : {}),
    ...(block.hiddenBookmarks ? { hiddenBookmarks: block.hiddenBookmarks } : {}),
    ...(block.commentStarts ? { commentStarts: block.commentStarts } : {}),
    ...(block.commentEnds ? { commentEnds: block.commentEnds } : {}),
    ...(block.sdtShell ? { sdtShell: block.sdtShell } : {}),
    runs,
  }
}

/**
 * The block-level w:ins / w:del wrapper of a whole block, unless the block
 * sits in a content control: the shell bytes ride on the block, and a wrapper
 * around them would open w:del before w:sdt. The runs carry the marks there.
 */
function blockRevision(
  block: Block,
  kind: 'ins' | 'del',
  author: string,
  date: string,
  ids: RevisionIds,
): { revision?: SaveBlock['revision'] } {
  if (block.sdtShell) return {}
  return { revision: { kind, author, date, id: ids.take() } }
}

/** an unchanged block: the original bytes, untouched */
function originalPassthrough(block: Block): SaveBlock {
  if (block.docxIndex !== null) return { kind: 'original', docxIndex: block.docxIndex }
  return { kind: 'generated', block: generatedFrom(block, block.runs ?? []) }
}

/** a block the revised document dropped */
function deletedBlock(block: Block, author: string, date: string, ids: RevisionIds): SaveBlock {
  const revision: RevisionInfo = { author, date, id: ids.take() }
  if (block.runs) {
    const runs = comparableRuns(block.runs).map((r) => ({ ...foreignRun(r), del: revision }))
    return {
      kind: 'generated',
      block: generatedFrom(block, runs),
      ...blockRevision(block, 'del', author, date, ids),
    }
  }
  if (block.type === 'image') {
    return {
      kind: 'generated',
      block: generatedFrom(block, [{ text: '[picture]', del: revision }]),
      revision: { kind: 'del', author, date, id: ids.take() },
    }
  }
  return {
    kind: 'xml',
    xml: withoutForeignObjects(block.originalXml ?? ''),
    revision: { kind: 'del', author, date, id: ids.take() },
  }
}

/** a block the revised document added */
function insertedBlock(block: Block, author: string, date: string, ids: RevisionIds): SaveBlock {
  const revision: RevisionInfo = { author, date, id: ids.take() }
  if (block.runs) {
    const runs = comparableRuns(block.runs).map((r) => ({ ...emittedRun(r), ins: revision }))
    return {
      kind: 'generated',
      block: generatedFrom(block, runs),
      ...blockRevision(block, 'ins', author, date, ids),
    }
  }
  if (block.docxIndex !== null) {
    return {
      kind: 'original',
      docxIndex: block.docxIndex,
      revision: { kind: 'ins', author, date, id: ids.take() },
    }
  }
  return {
    kind: 'xml',
    xml: block.originalXml ?? '',
    revision: { kind: 'ins', author, date, id: ids.take() },
  }
}

/** stands in for a picture in the compared text, so a picture is a token the diff keeps or moves */
const PICTURE_TOKEN = '\uFFFC'

/**
 * The runs as a reader sees them with every pending change accepted: deleted
 * text is gone, inserted text is plain. A picture becomes a one-character
 * token so it survives slicing by character range.
 */
function comparableRuns(runs: Run[]): Run[] {
  const out: Run[] = []
  for (const run of runs) {
    if (run.del) continue
    const { ins: _ins, ...rest } = run
    out.push(rest.image && !rest.text ? { ...rest, text: PICTURE_TOKEN } : rest)
  }
  return out
}

/** a run of the revised document, the picture token folded back into the picture run */
function emittedRun(run: Run): Run {
  return run.image && run.text === PICTURE_TOKEN ? { ...run, text: '' } : run
}

/**
 * A run of the original document that ends up in the redline: its picture
 * points at a media part the revised package does not have, so the deletion
 * names it instead of carrying it.
 */
function foreignRun(run: Run): Run {
  if (!run.image) return run
  const { image: _image, ...rest } = run
  return { ...rest, text: '[picture]' }
}

/** drop the drawings, VML pictures and embedded objects of a fragment from the other package */
function withoutForeignObjects(xml: string): string {
  return xml
    .replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, '')
    .replace(/<w:pict(?:\s[^>]*)?>[\s\S]*?<\/w:pict>/g, '')
    .replace(/<w:object(?:\s[^>]*)?>[\s\S]*?<\/w:object>/g, '')
}

/**
 * Revision ids are unique across the document: they start above every id the
 * revised document already uses, so the redline never repeats one.
 */
class RevisionIds {
  private next: number

  constructor(documentXml: string) {
    let max = 0
    for (const m of documentXml.matchAll(/\bw:id\s*=\s*["'](\d+)["']/g)) {
      max = Math.max(max, parseInt(m[1]!, 10))
    }
    this.next = Math.max(9001, max + 1)
  }

  take(): string {
    return String(this.next++)
  }

  /** the first id not handed out yet */
  peek(): number {
    return this.next
  }
}

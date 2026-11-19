/**
 * Compare two documents the way Word's Review > Compare does: the revised
 * document's content with the differences marked as tracked insertions and
 * deletions, so the result can be accepted or rejected like any other redline.
 *
 * The revision marks are the ones the rest of the suite already writes. A run
 * carries `ins` / `del` (see Run in types.ts) and the generator wraps
 * consecutive runs of the same revision in w:ins / w:del (see generate.ts);
 * a paragraph inserted or deleted as a whole additionally carries the record
 * on its paragraph mark (w:pPr/w:rPr) — Word's own Compare shape, which
 * accept / reject resolve to the revised / original document. A non-paragraph
 * block (a table) keeps the block-level w:ins / w:del wrapper. A redline
 * written here therefore resolves through the same accept / reject path.
 *
 * Text whose formatting alone changed carries a w:rPrChange with the old
 * formatting, and a body paragraph whose own properties changed a
 * w:pPrChange with the original's. A paragraph that sits elsewhere with the
 * same text is a move (w:moveFrom / w:moveTo). Headers, footers, footnotes,
 * endnotes, text boxes and the cells of a table that kept its shape are
 * compared as text (compare-story.ts).
 *
 * Comparing reads each document with its pending changes accepted. Combining
 * (`keepRevisions`) reads the text with its markup: a change either document
 * already tracks stays under its own author and date, and only what differs
 * beyond those is a new change.
 *
 * Out of scope: moves shorter than a paragraph, pictures and fields inside a
 * changed header, footer, note or table cell.
 */
import type JSZip from 'jszip'
import {
  childElements,
  diffStoryXml,
  diffTableXml,
  diffTextBoxXml,
  insertedStoryXml,
  renumberedParaMark,
  type StoryDiffContext,
} from './compare-story'
import {
  existingPPrChange,
  paragraphFormatKey,
  paragraphWithPPr,
  recordedPPr,
  withPPrChange,
} from './compare-ppr'
import { paraMarkRevisionXml, setParaMarkRevision } from './generate'
import { NOTE_PART_PATH, type NoteKind } from './notes'
import { parseRels, resolveMainDocumentPath, resolveRelationshipTargetPath } from './parse-package'
import { saveDocx, type ParsedDocFull, type SaveBlock } from './patch'
import { hfReferenceRId, hfReferenceTags, hfReferenceType } from './section'
import type { Block, GeneratedBlock, RevisionInfo, Run } from './types'
import { escapeXmlAttr } from './xml-utils'
import { loadDocxZip } from './zip-load'

/** at what granularity a changed paragraph is diffed */
export type CompareGranularity = 'word' | 'char'

/** what the comparison leaves out */
export type CompareIgnore = 'formatting' | 'case' | 'whitespace'

/** what the comparison covers; everything is on unless switched off */
export interface CompareSettings {
  /** a paragraph that sits elsewhere with the same text is a move, not a deletion and an insertion */
  moves: boolean
  /** the result carries the revised document's comments */
  comments: boolean
  /** text whose formatting alone changed is a tracked formatting change */
  formatting: boolean
  caseChanges: boolean
  whiteSpace: boolean
  tables: boolean
  headersFooters: boolean
  /** footnotes and endnotes */
  notes: boolean
  /** the text inside a text box is compared; off leaves a text box as the revised document has it */
  textBoxes: boolean
  /** a field whose code changed is a change even when it shows the same text; off sets every difference inside a field aside */
  fields: boolean
}

export const DEFAULT_COMPARE_SETTINGS: CompareSettings = {
  moves: true,
  comments: true,
  formatting: true,
  caseChanges: true,
  whiteSpace: true,
  tables: true,
  headersFooters: true,
  notes: true,
  textBoxes: true,
  fields: true,
}

export interface CompareOptions {
  /** author recorded on every revision; default "GenOffice" */
  author?: string
  /** ISO timestamp recorded on every revision; default: now */
  date?: string
  /** default "word" */
  granularity?: CompareGranularity
  /** default: nothing is ignored */
  ignore?: readonly CompareIgnore[]
  settings?: Partial<CompareSettings>
  /**
   * Combine: tracked changes either document already holds stay in the result
   * under their own authors and dates. Off, both are read as accepted.
   */
  keepRevisions?: boolean
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
  /** text regions (w:rPrChange) and paragraphs (w:pPrChange) whose formatting alone changed */
  formatChanges: number
  /** tracked changes of the original document kept under their own author (keepRevisions) */
  carried: number
  /** paragraphs that moved (a w:moveFrom / w:moveTo pair each) */
  moves: number
  /** differing paragraphs and tables left as the revised document has them */
  skipped: number
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
  /** differences inside a field are not differences */
  ignoreFields?: boolean
  /** a kept field whose code differs is replaced */
  fieldCodes?: boolean
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

type Aligned<U = Unit> =
  { kind: 'pair'; left: U; right: U } | { kind: 'removed'; left: U } | { kind: 'added'; right: U }

/** a piece of the compared paragraph: text kept, deleted or inserted */
interface Chunk {
  kind: 'keep' | 'del' | 'ins' | 'fmt'
  /** char range in the original text */
  left?: [number, number]
  /** char range in the revised text */
  right?: [number, number]
  /** fmt: the original run whose formatting the text had */
  leftRun?: number
  /** keep: the original text is tracked here, character for character over the revised range */
  carry?: boolean
}

/**
 * How a character stands in its document's tracked changes: plain (''), part
 * of a pending deletion ('del': still text of the document as it was), or
 * pending inserted text ('ins': not there before the change).
 */
export type TrackedState = '' | 'del' | 'ins'

/** the tracked-change state of every character of each side */
export interface TrackedFlags {
  left: readonly TrackedState[]
  right: readonly TrackedState[]
  /** the side whose changes the other takes on where both have the text; default the original */
  carryFrom?: 'left' | 'right'
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
  const setup = setupOf(options)
  const leftZip = await loadDocxZip(original.internal.originalBytes)
  const rightZip = await loadDocxZip(revised.internal.originalBytes)
  const stories = [
    ...(setup.settings.headersFooters ? await headerFooterStories(leftZip, rightZip) : []),
    ...(setup.settings.notes ? await noteStories(leftZip, rightZip) : []),
  ]
  const ids = new RevisionIds([revised.internal.documentXml, ...stories.map((s) => s.right)])
  const plan = planBody(original, revised, setup, ids)
  const partXml: Record<string, string> = {}
  const ctx = storyContext(setup, ids, plan.summary)
  for (const story of stories) {
    const xml = story.diff(ctx)
    if (xml !== story.right) partXml[story.path] = xml
  }
  const bytes = await saveDocx(revised, plan.blocks, {
    revisionIdFloor: ids.peek(),
    partXml,
    ...(setup.settings.comments ? {} : { comments: [] }),
  })
  return { bytes, summary: plan.summary }
}

interface Setup {
  rules: Rules
  settings: CompareSettings
  author: string
  date: string
  granularity: CompareGranularity
  keepRevisions: boolean
}

function setupOf(options: CompareOptions): Setup {
  const settings = { ...DEFAULT_COMPARE_SETTINGS, ...options.settings }
  return {
    settings,
    rules: {
      ignoreCase: !settings.caseChanges || (options.ignore?.includes('case') ?? false),
      ignoreWhitespace: !settings.whiteSpace || (options.ignore?.includes('whitespace') ?? false),
      ignoreFormatting: !settings.formatting || (options.ignore?.includes('formatting') ?? false),
      ignoreFields: !settings.fields,
      fieldCodes: settings.fields,
    },
    author: options.author?.trim() || DEFAULT_COMPARE_AUTHOR,
    date: options.date || new Date().toISOString(),
    granularity: options.granularity === 'char' ? 'char' : 'word',
    keepRevisions: options.keepRevisions === true,
  }
}

/** the save plan of the body comparison, for a caller that writes the file itself */
export function compareBlocks(
  original: ParsedDocFull,
  revised: ParsedDocFull,
  options: CompareOptions = {},
): { blocks: SaveBlock[]; summary: CompareSummary; nextRevisionId: number } {
  const ids = new RevisionIds([revised.internal.documentXml])
  const plan = planBody(original, revised, setupOf(options), ids)
  return { ...plan, nextRevisionId: ids.peek() }
}

function planBody(
  original: ParsedDocFull,
  revised: ParsedDocFull,
  setup: Setup,
  ids: RevisionIds,
): { blocks: SaveBlock[]; summary: CompareSummary } {
  const { rules, settings, author, date, granularity, keepRevisions } = setup

  const left = unitsOf(original, keepRevisions, false)
  const right = unitsOf(revised, keepRevisions, keepRevisions)
  const aligned = alignUnits(left, right, rules)

  const blocks: SaveBlock[] = []
  const summary: CompareSummary = {
    insertedParagraphs: 0,
    deletedParagraphs: 0,
    changedParagraphs: 0,
    unchangedParagraphs: 0,
    insertions: 0,
    deletions: 0,
    formatChanges: 0,
    carried: 0,
    moves: 0,
    skipped: 0,
  }
  const moved = settings.moves
    ? movedParagraphs(aligned, rules, keepRevisions)
    : new Map<Aligned, string>()
  const oldStyle = (run: Run) => oldFormatting(run, revised)
  const readRuns = keepRevisions ? markupRuns : comparableRuns
  const carry = carrier(ids, summary)
  const formats = paragraphFormats(original, revised)

  for (const entry of aligned) {
    if (!settings.tables && (entry.kind === 'removed' ? entry.left : entry.right).block.table) {
      if (entry.kind !== 'removed') blocks.push(originalPassthrough(entry.right.block))
      continue
    }
    const move = moved.get(entry)
    if (move) {
      if (entry.kind === 'removed') {
        summary.moves++
        blocks.push(movedBlock(entry.left.block, 'del', move, author, date, ids))
      } else if (entry.kind === 'added') {
        blocks.push(movedBlock(entry.right.block, 'ins', move, author, date, ids))
      }
      continue
    }
    if (entry.kind === 'pair') {
      const leftBlock = entry.left.block
      const rightBlock = entry.right.block
      if (keepRevisions && rightBlock.blockRevision) {
        blocks.push(originalPassthrough(rightBlock))
        continue
      }
      if (leftBlock.runs && rightBlock.runs) {
        const leftRuns = readRuns(leftBlock.runs)
        const rightRuns = readRuns(rightBlock.runs)
        const chunks = paragraphChunks(
          leftRuns,
          rightRuns,
          granularity,
          rules,
          keepRevisions
            ? { left: trackedFlags(leftRuns), right: trackedFlags(rightRuns) }
            : undefined,
        )
        let rawPPr = rules.ignoreFormatting
          ? null
          : formats.changed(leftBlock, rightBlock, keepRevisions, { author, date }, ids)
        if (rawPPr !== null) {
          if (rawPPr.carried) summary.carried++
          else summary.formatChanges++
        }
        if (keepRevisions) {
          for (const kind of ['del', 'ins'] as const) {
            const key = kind === 'del' ? 'paraMarkDel' : 'paraMarkIns'
            const mark = leftBlock[key]
            if (!mark || rightBlock[key]) continue
            summary.carried++
            const record = JSON.stringify({ author: mark.author, date: mark.date })
            rawPPr = {
              xml: setParaMarkRevision(
                rawPPr?.xml ?? rightBlock.rawPPr ?? '',
                paraMarkRevisionXml(kind, record, () => ids.take()),
              ),
            }
          }
        }
        if (!hasRevision(chunks)) {
          if (rawPPr === null) {
            summary.unchangedParagraphs++
            blocks.push(originalPassthrough(rightBlock))
            continue
          }
          summary.changedParagraphs++
          const xml = rightBlock.sdtShell
            ? null
            : paragraphWithPPr(rightBlock.originalXml, rightBlock.rawPPr, rawPPr.xml)
          blocks.push(
            xml !== null
              ? {
                  kind: 'xml',
                  xml,
                  ...(rightBlock.docxIndex !== null ? { docxIndex: rightBlock.docxIndex } : {}),
                }
              : {
                  kind: 'generated',
                  block: {
                    ...generatedFrom(rightBlock, rightRuns.map(emittedRun)),
                    rawPPr: rawPPr.xml,
                  },
                },
          )
          continue
        }
        summary.changedParagraphs++
        const runs: Run[] = []
        for (const chunk of chunks) {
          if (chunk.kind === 'keep') {
            const kept = sliceRuns(rightRuns, chunk.right![0], chunk.right![1]).map(emittedRun)
            if (!chunk.carry) {
              runs.push(...kept)
              continue
            }
            // the original tracks this text and the revised document has it plain
            let at = chunk.right![0]
            for (const mine of sliceRuns(leftRuns, chunk.left![0], chunk.left![1])) {
              const to = at + mine.text.length
              for (const r of sliceRuns(rightRuns, at, to).map(emittedRun))
                runs.push(isTracked(r) || !isTracked(mine) ? r : { ...r, ...carry.marks(mine) })
              at = to
            }
            continue
          }
          if (chunk.kind === 'fmt') {
            const old = oldStyle(leftRuns[chunk.leftRun!]!)
            for (const r of sliceRuns(rightRuns, chunk.right![0], chunk.right![1])) {
              if (keepRevisions && r.rPrChange) {
                runs.push(emittedRun(r))
                continue
              }
              summary.formatChanges++
              runs.push({ ...emittedRun(r), rPrChange: { author, date, id: ids.take(), old } })
            }
            continue
          }
          const info: RevisionInfo = { author, date, id: ids.take() }
          if (chunk.kind === 'del') {
            const gone = foreignRuns(sliceRuns(leftRuns, chunk.left![0], chunk.left![1]))
            if (gone.some((r) => !isTracked(r))) summary.deletions++
            runs.push(
              ...gone.map((r) =>
                isTracked(r) ? { ...r, ...carry.marks(r) } : { ...r, del: info },
              ),
            )
          } else {
            const added = sliceRuns(rightRuns, chunk.right![0], chunk.right![1]).map(emittedRun)
            if (added.some((r) => !isTracked(r))) summary.insertions++
            runs.push(...added.map((r) => (isTracked(r) ? r : { ...r, ins: info })))
          }
        }
        blocks.push({
          kind: 'generated',
          block: {
            ...generatedFrom(rightBlock, runs),
            ...(rawPPr !== null ? { rawPPr: rawPPr.xml } : {}),
          },
        })
        continue
      }
      // a table, a picture or a field: the whole block is kept or replaced
      if (comparableText(entry.left.text, rules) === comparableText(entry.right.text, rules)) {
        summary.unchangedParagraphs++
        blocks.push(originalPassthrough(rightBlock))
        continue
      }
      if (leftBlock.textboxes && rightBlock.textboxes) {
        if (!settings.textBoxes) {
          summary.unchangedParagraphs++
          blocks.push(originalPassthrough(rightBlock))
          continue
        }
        const xml = diffTextBoxXml(
          leftBlock.originalXml ?? '',
          rightBlock.originalXml ?? '',
          storyContext(setup, ids, summary),
        )
        if (xml !== null) {
          summary.changedParagraphs++
          blocks.push({
            kind: 'xml',
            xml,
            ...(rightBlock.docxIndex !== null ? { docxIndex: rightBlock.docxIndex } : {}),
          })
          continue
        }
      }
      if (leftBlock.table && rightBlock.table && leftBlock.originalXml && rightBlock.originalXml) {
        const xml = diffTableXml(
          leftBlock.originalXml,
          rightBlock.originalXml,
          storyContext(setup, ids, summary),
        )
        if (xml !== null) {
          summary.changedParagraphs++
          blocks.push({
            kind: 'xml',
            xml,
            ...(rightBlock.docxIndex !== null ? { docxIndex: rightBlock.docxIndex } : {}),
          })
          continue
        }
      }
      summary.deletedParagraphs++
      summary.insertedParagraphs++
      summary.deletions++
      summary.insertions++
      blocks.push(deletedBlock(leftBlock, author, date, ids))
      blocks.push(insertedBlock(rightBlock, author, date, ids))
      continue
    }
    if (entry.kind === 'removed') {
      const gone = entry.left.block
      if (keepRevisions && gone.runs && whollyTracked(gone)) {
        // a paragraph the original inserted or deleted as a tracked change
        const block = generatedFrom(
          gone,
          foreignRuns(markupRuns(gone.runs)).map((r) => ({ ...r, ...carry.marks(r) })),
        )
        if (block.rawPPr) block.rawPPr = renumberedParaMark(block.rawPPr, () => ids.take())
        if (!block.runs.some(isTracked)) summary.carried++
        blocks.push({ kind: 'generated', block })
        continue
      }
      summary.deletedParagraphs++
      summary.deletions++
      blocks.push(deletedBlock(gone, author, date, ids, keepRevisions ? carry : undefined))
      continue
    }
    const added = entry.right.block
    if (keepRevisions && (added.runs ? whollyTracked(added) : !!added.blockRevision)) {
      blocks.push(originalPassthrough(added))
      continue
    }
    summary.insertedParagraphs++
    summary.insertions++
    blocks.push(insertedBlock(added, author, date, ids, keepRevisions))
  }

  return { blocks, summary }
}

/** the text diff of a story outside the body model, under this comparison's rules */
function storyContext(setup: Setup, ids: RevisionIds, counts: CompareSummary): StoryDiffContext {
  const rules: Rules = { ...setup.rules, ignoreFormatting: true }
  const units = (texts: string[]) => texts.map((text, index) => ({ index, text }))
  return {
    author: setup.author,
    date: setup.date,
    takeId: () => ids.take(),
    align: (left, right) => alignmentOf(alignUnits(units(left), units(right), rules)),
    chunks: (left, right, tracked) =>
      paragraphChunks([{ text: left }], [{ text: right }], setup.granularity, rules, tracked),
    keepRevisions: setup.keepRevisions,
    same: (left, right) => comparableText(left, rules) === comparableText(right, rules),
    tables: setup.settings.tables,
    counts,
  }
}

const alignmentOf = (aligned: Array<Aligned<{ index: number }>>): TextAlignment[] =>
  aligned.map((entry) =>
    entry.kind === 'pair'
      ? { kind: 'pair', left: entry.left.index, right: entry.right.index }
      : entry.kind === 'removed'
        ? { kind: 'removed', left: entry.left.index }
        : { kind: 'added', right: entry.right.index },
  )

interface Story {
  path: string
  /** the revised part as it is */
  right: string
  diff: (ctx: StoryDiffContext) => string
}

const partText = async (zip: JSZip, path: string): Promise<string | null> =>
  (await zip.file(path)?.async('string')) ?? null

/** the header and footer parts of a package, keyed by kind, page type and how many came before */
async function headerFooterParts(zip: JSZip): Promise<Map<string, string>> {
  const out = new Map<string, string>()
  const main = await resolveMainDocumentPath(zip)
  const documentXml = main ? await partText(zip, main) : null
  if (!main || !documentXml) return out
  const slash = main.lastIndexOf('/')
  const rels = await parseRels(
    zip,
    `${main.slice(0, slash + 1)}_rels/${main.slice(slash + 1)}.rels`,
  )
  const seen = new Map<string, number>()
  const used = new Set<string>()
  for (const sectPr of documentXml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g) ?? []) {
    for (const kind of ['header', 'footer'] as const) {
      for (const tag of hfReferenceTags(sectPr, kind)) {
        const target = rels.get(hfReferenceRId(tag) ?? '')?.target
        const path = target ? resolveRelationshipTargetPath(main, target) : null
        if (!path || used.has(path)) continue
        used.add(path)
        const slot = `${kind}:${hfReferenceType(tag) ?? 'default'}`
        const nth = seen.get(slot) ?? 0
        seen.set(slot, nth + 1)
        out.set(`${slot}:${nth}`, path)
      }
    }
  }
  return out
}

const STORY_ROOT = /^([\s\S]*?<w:(hdr|ftr)\b[^>]*>)([\s\S]*)(<\/w:\2>\s*)$/

async function headerFooterStories(leftZip: JSZip, rightZip: JSZip): Promise<Story[]> {
  const mine = await headerFooterParts(leftZip)
  const out: Story[] = []
  for (const [key, path] of await headerFooterParts(rightZip)) {
    const right = await partText(rightZip, path)
    const parts = right ? STORY_ROOT.exec(right) : null
    if (!right || !parts) continue
    const twin = mine.get(key)
    const left = twin ? STORY_ROOT.exec((await partText(leftZip, twin)) ?? '') : null
    out.push({
      path,
      right,
      diff: (ctx) =>
        parts[1] +
        (left ? diffStoryXml(left[3]!, parts[3]!, ctx) : insertedStoryXml(parts[3]!, ctx)) +
        parts[4],
    })
  }
  return out
}

/** the notes of a part that are text, not separators */
function noteEntries(xml: string, kind: NoteKind): Array<{ inner: string; at: number }> {
  const out: Array<{ inner: string; at: number }> = []
  const re = new RegExp(`<w:${kind}\\b([^>]*)>([\\s\\S]*?)<\\/w:${kind}>`, 'g')
  for (const m of xml.matchAll(re)) {
    const type = /\bw:type\s*=\s*["']([^"']*)["']/.exec(m[1]!)?.[1]
    if (type && type !== 'normal') continue
    out.push({ inner: m[2]!, at: m.index + m[0].indexOf('>') + 1 })
  }
  return out
}

const noteText = (inner: string): string =>
  childElements(inner)
    .map((element) => element.xml.replace(/<[^>]+>/g, ''))
    .join('\n')

async function noteStories(leftZip: JSZip, rightZip: JSZip): Promise<Story[]> {
  const out: Story[] = []
  for (const kind of ['footnote', 'endnote'] as const) {
    const path = NOTE_PART_PATH[kind]
    const right = await partText(rightZip, path)
    if (!right) continue
    const theirs = noteEntries(right, kind)
    const mine = noteEntries((await partText(leftZip, path)) ?? '', kind)
    out.push({
      path,
      right,
      diff: (ctx) => {
        const twins = new Map<number, number>()
        for (const entry of ctx.align(
          mine.map((n) => noteText(n.inner)),
          theirs.map((n) => noteText(n.inner)),
        ))
          if (entry.kind === 'pair') twins.set(entry.right, entry.left)
        let xml = ''
        let at = 0
        theirs.forEach((note, i) => {
          const twin = twins.get(i)
          xml += right.slice(at, note.at)
          xml +=
            twin === undefined
              ? insertedStoryXml(note.inner, ctx)
              : diffStoryXml(mine[twin]!.inner, note.inner, ctx)
          at = note.at + note.inner.length
        })
        return xml + right.slice(at)
      },
    })
  }
  return out
}

/**
 * Paragraphs only the original has whose text a paragraph only the revised
 * document has repeats: the paragraph moved. Each pair shares a move name.
 */
function movedParagraphs(
  aligned: Aligned[],
  rules: Rules,
  keepRevisions: boolean,
): Map<Aligned, string> {
  const movable = (unit: Unit): boolean =>
    unit.textual &&
    !(keepRevisions && unit.block.runs?.some(isTracked)) &&
    !unit.block.sdtShell &&
    !unit.text.includes(PICTURE_TOKEN) &&
    /\S/.test(unit.text)
  const gone = new Map<string, Aligned[]>()
  for (const entry of aligned) {
    if (entry.kind !== 'removed' || !movable(entry.left)) continue
    const key = comparableText(entry.left.text, rules)
    gone.set(key, [...(gone.get(key) ?? []), entry])
  }
  const out = new Map<Aligned, string>()
  for (const entry of aligned) {
    if (entry.kind !== 'added' || !movable(entry.right)) continue
    const source = gone.get(comparableText(entry.right.text, rules))?.shift()
    if (!source) continue
    const name = `move${out.size / 2 + 1}`
    out.set(source, name)
    out.set(entry, name)
  }
  return out
}

/** a paragraph that moved, at its old place (del) or its new one (ins) */
function movedBlock(
  block: Block,
  kind: 'ins' | 'del',
  move: string,
  author: string,
  date: string,
  ids: RevisionIds,
): SaveBlock {
  const info: RevisionInfo = { author, date, id: ids.take(), move }
  const source = comparableRuns(block.runs ?? [])
  const runs =
    kind === 'del'
      ? foreignRuns(source).map((r) => ({ ...r, del: info }))
      : source.map((r) => ({ ...emittedRun(r), ins: info }))
  const mark = JSON.stringify({ author, date })
  return {
    kind: 'generated',
    block: {
      ...generatedFrom(block, runs),
      ...(kind === 'del' ? { paraMarkDel: mark } : { paraMarkIns: mark }),
    },
  }
}

const isTracked = (run: Run): boolean => !!(run.ins || run.del)

const trackedFlags = (runs: Run[]): TrackedState[] =>
  runs.flatMap((run) =>
    new Array<TrackedState>(run.text.length).fill(run.ins ? 'ins' : run.del ? 'del' : ''),
  )

/** a paragraph that is one tracked change: every run inserted or deleted, or an empty one whose mark is */
function whollyTracked(block: Block): boolean {
  const runs = (block.runs ?? []).filter((run) => run.text !== '' || run.image)
  return runs.length ? runs.every(isTracked) : !!(block.paraMarkIns || block.paraMarkDel)
}

interface Carrier {
  /** the tracked-change marks of a run of the original, under ids of this document */
  marks: (run: Run) => Pick<Run, 'ins' | 'del'>
}

/**
 * Tracked changes of the original keep their author and date and take a new
 * id; a move loses its name, since its other half may not come along.
 */
function carrier(ids: RevisionIds, summary: CompareSummary): Carrier {
  const seen = new Map<string, RevisionInfo>()
  const one = (kind: 'ins' | 'del', info: RevisionInfo): RevisionInfo => {
    const key = JSON.stringify([kind, info.author, info.date ?? null, info.id ?? null])
    let out = seen.get(key)
    if (!out) {
      const { move: _move, ...rest } = info
      out = { ...rest, id: ids.take() }
      seen.set(key, out)
      summary.carried++
    }
    return out
  }
  return {
    marks: (run) => ({
      ...(run.ins ? { ins: one('ins', run.ins) } : {}),
      ...(run.del ? { del: one('del', run.del) } : {}),
    }),
  }
}

/** a paragraph's own properties with the tracked change the comparison adds; null when it adds none */
function paragraphFormats(original: ParsedDocFull, revised: ParsedDocFull) {
  const listKey = (doc: ParsedDocFull) => (numId: string, ilvl: number) =>
    JSON.stringify(doc.numbering.get(numId)?.levels[ilvl] ?? null)
  const leftList = listKey(original)
  const rightList = listKey(revised)
  const keep = {
    style: (styleId: string) => revised.styles.has(styleId),
    list: (numId: string, ilvl: number) =>
      revised.numbering.has(numId) && leftList(numId, ilvl) === rightList(numId, ilvl),
  }
  const attrs = (id: string, author: string, date?: string) =>
    ` w:id="${id}" w:author="${escapeXmlAttr(author)}"` +
    (date ? ` w:date="${escapeXmlAttr(date)}"` : '')
  return {
    changed(
      left: Block,
      right: Block,
      keepRevisions: boolean,
      label: { author: string; date: string },
      ids: RevisionIds,
    ): { xml: string; carried?: boolean } | null {
      if (keepRevisions && right.pPrChangeInfo) return null
      const same =
        paragraphFormatKey(left.rawPPr, leftList) === paragraphFormatKey(right.rawPPr, rightList)
      if (!same) {
        return {
          xml: withPPrChange(
            right.rawPPr,
            recordedPPr(left.rawPPr, keep),
            attrs(ids.take(), label.author, label.date),
          ),
        }
      }
      // the original tracks the formatting both now have
      const theirs = keepRevisions ? existingPPrChange(left.rawPPr) : null
      if (!theirs) return null
      return {
        xml: withPPrChange(
          right.rawPPr,
          recordedPPr(theirs.recorded, keep),
          attrs(ids.take(), theirs.author, theirs.date),
        ),
        carried: true,
      }
    },
  }
}

/** what a run of the original looked like, as the record of a formatting change */
function oldFormatting(run: Run, revised: ParsedDocFull): NonNullable<Run['rPrChange']>['old'] {
  return {
    ...(run.bold ? { bold: true } : {}),
    ...(run.italic ? { italic: true } : {}),
    ...(run.underline ? { underline: true } : {}),
    ...(run.strike ? { strike: true } : {}),
    ...(run.dstrike ? { dstrike: true } : {}),
    ...(run.color ? { color: run.color } : {}),
    ...(run.sizeHalfPoints ? { sizeHalfPoints: run.sizeHalfPoints } : {}),
    ...(run.font ? { font: run.font } : {}),
    ...(run.fontAscii ? { fontAscii: run.fontAscii } : {}),
    ...(run.charSpacingTwips ? { charSpacingTwips: run.charSpacingTwips } : {}),
    ...(run.charScalePct ? { charScalePct: run.charScalePct } : {}),
    ...(run.highlight ? { highlight: run.highlight } : {}),
    ...(run.vertAlign ? { vertAlign: run.vertAlign } : {}),
    ...(run.caps ? { caps: run.caps } : {}),
    ...(run.shading ? { shading: run.shading } : {}),
    // a character style only the original defines would name nothing here
    ...(run.styleId && revised.styles.has(run.styleId) ? { styleId: run.styleId } : {}),
  }
}

/** how two lists of block texts line up: a block both have, or one only one side has */
export type TextAlignment =
  | { kind: 'pair'; left: number; right: number }
  | { kind: 'removed'; left: number }
  | { kind: 'added'; right: number }

/**
 * The paragraph alignment of a comparison, for a caller that holds the two
 * documents in its own model: identical texts anchor the sequence and the
 * texts between two anchors pair up when they share enough words.
 */
export function alignTexts(
  left: readonly string[],
  right: readonly string[],
  ignore: TextCompareIgnore = {},
): TextAlignment[] {
  const rules = textRules(ignore)
  const units = (texts: readonly string[]) => texts.map((text, index) => ({ index, text }))
  return alignmentOf(alignUnits(units(left), units(right), rules))
}

/** a piece of a compared text: kept, deleted (a range of the original) or inserted (a range of the revised) */
export interface TextChunk {
  kind: 'keep' | 'del' | 'ins'
  left?: [number, number]
  right?: [number, number]
}

/** differences a text comparison sets aside */
export interface TextCompareIgnore {
  case?: boolean
  whitespace?: boolean
}

const textRules = (ignore: TextCompareIgnore): Rules => ({
  ignoreCase: ignore.case ?? false,
  ignoreWhitespace: ignore.whitespace ?? false,
  ignoreFormatting: true,
})

/** the word-level (or character-level) differences of two paragraph texts, formatting set aside */
export function diffText(
  left: string,
  right: string,
  granularity: CompareGranularity = 'word',
  ignore: TextCompareIgnore = {},
): TextChunk[] {
  return paragraphChunks([{ text: left }], [{ text: right }], granularity, textRules(ignore)).map(
    ({ kind, left: l, right: r }) => ({
      kind: kind === 'fmt' ? 'keep' : kind,
      ...(l ? { left: l } : {}),
      ...(r ? { right: r } : {}),
    }),
  )
}

/**
 * The differences of two paragraph texts that hold tracked changes, both read
 * with their markup: a `keep` chunk flagged `carry` is text one side tracks,
 * character for character over the other side's range.
 */
export function diffTrackedText(
  left: string,
  right: string,
  tracked: TrackedFlags,
  granularity: CompareGranularity = 'word',
  ignore: TextCompareIgnore = {},
): Array<TextChunk & { carry?: boolean }> {
  return paragraphChunks(
    [{ text: left }],
    [{ text: right }],
    granularity,
    textRules(ignore),
    tracked,
  ).map(({ kind, left: l, right: r, carry }) => ({
    kind: kind === 'fmt' ? 'keep' : kind,
    ...(l ? { left: l } : {}),
    ...(r ? { right: r } : {}),
    ...(carry ? { carry } : {}),
  }))
}

/** the comparable units of a document, in body order */
function unitsOf(parsed: ParsedDocFull, markup: boolean, keepDeleted: boolean): Unit[] {
  const out: Unit[] = []
  parsed.blocks.forEach((block, index) => {
    if (block.hidden || (!keepDeleted && block.blockRevision?.kind === 'del')) return
    const textual = Array.isArray(block.runs)
    out.push({ index, block, textual, text: contentSignature(block, markup) })
  })
  return out
}

/** the text a block is compared by: its runs, a table's cell texts, or its preview */
function contentSignature(block: Block, markup: boolean): string {
  if (block.runs)
    return (markup ? markupRuns : comparableRuns)(block.runs)
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
function alignUnits<U extends { text: string }>(left: U[], right: U[], rules: Rules): Aligned<U>[] {
  const anchors = anchorPairs(
    left.map((u) => comparableText(u.text, rules)),
    right.map((u) => comparableText(u.text, rules)),
  )
  const out: Aligned<U>[] = []
  let li = 0
  let ri = 0
  for (const [l, r] of anchors) {
    out.push(...pairRegion(left.slice(li, l), right.slice(ri, r), rules))
    out.push({ kind: 'pair', left: left[l]!, right: right[r]! })
    li = l + 1
    ri = r + 1
  }
  out.push(...pairRegion(left.slice(li), right.slice(ri), rules))
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
function pairRegion<U extends { text: string }>(left: U[], right: U[], rules: Rules): Aligned<U>[] {
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
  const score = (a: U, b: U): number =>
    similarity(comparableText(a.text, rules), comparableText(b.text, rules))
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
  const out: Aligned<U>[] = []
  let i = n
  let j = m
  const region: Aligned<U>[] = []
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
  tracked?: TrackedFlags,
): Chunk[] {
  const leftText = leftRuns.map((r) => r.text).join('')
  const rightText = rightRuns.map((r) => r.text).join('')
  const ops = tracked
    ? trackedTokenOps(leftText, rightText, granularity, tracked, rules)
    : tokenOps(tokenize(leftText, granularity), tokenize(rightText, granularity), rules)

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
  const leftField = (t: Token) => fieldCodeOf(leftRuns, leftStarts, t)
  const rightField = (t: Token) => fieldCodeOf(rightRuns, rightStarts, t)
  for (const op of ops) {
    if (op.kind === 'del') {
      // a whitespace run the caller asked to ignore is not a difference
      if (rules.ignoreWhitespace && !op.l.text.trim()) continue
      if (rules.ignoreFields && leftField(op.l) !== null) continue
      flush()
      chunks.push({ kind: 'del', left: [op.l.start, op.l.end] })
      continue
    }
    if (op.kind === 'ins') {
      if (rules.ignoreWhitespace && !op.r.text.trim()) continue
      flush()
      const inField = rules.ignoreFields && rightField(op.r) !== null
      chunks.push({ kind: inField ? 'keep' : 'ins', right: [op.r.start, op.r.end] })
      continue
    }
    if (rules.fieldCodes && leftField(op.l) !== rightField(op.r)) {
      flush()
      chunks.push({ kind: 'del', left: [op.l.start, op.l.end] })
      chunks.push({ kind: 'ins', right: [op.r.start, op.r.end] })
      continue
    }
    if (
      tracked &&
      op.l.end - op.l.start === op.r.end - op.r.start &&
      (tracked.carryFrom === 'right'
        ? tracked.right.slice(op.r.start, op.r.end)
        : tracked.left.slice(op.l.start, op.l.end)
      ).some((state) => state !== '')
    ) {
      flush()
      chunks.push({
        kind: 'keep',
        left: [op.l.start, op.l.end],
        right: [op.r.start, op.r.end],
        carry: true,
      })
      continue
    }
    const styled =
      !rules.ignoreFormatting &&
      styleOf(leftRuns, leftStarts, op.l) !== styleOf(rightRuns, rightStarts, op.r)
    if (styled) {
      flush()
      chunks.push({
        kind: 'fmt',
        left: [op.l.start, op.l.end],
        right: [op.r.start, op.r.end],
        leftRun: runIndexAt(leftStarts, op.l.start),
      })
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

type TokenOp =
  { kind: 'eq'; l: Token; r: Token } | { kind: 'del'; l: Token } | { kind: 'ins'; r: Token }

function tokenOps(leftTokens: Token[], rightTokens: Token[], rules: Rules): TokenOp[] {
  const pairs = lcsPairs(leftTokens, rightTokens, (a, b) => tokenEqual(a.text, b.text, rules))
  const ops: TokenOp[] = []
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
  return ops
}

type SideToken = Token & { fragment: boolean }

/**
 * The tokens of a text as it was before its pending insertions, placed in
 * the text as it stands. A token ends where a pending change starts or ends,
 * so each is one stretch of the text and wholly plain or wholly deleted.
 */
function baseTokens(
  text: string,
  states: readonly TrackedState[],
  granularity: CompareGranularity,
  from = 0,
  to = text.length,
): SideToken[] {
  const at: number[] = []
  let base = ''
  for (let i = from; i < to; i++) {
    if (states[i] === 'ins') continue
    at.push(i)
    base += text[i]
  }
  const word = (i: number) => i >= 0 && i < text.length && /\S/.test(text[i]!)
  const out: SideToken[] = []
  for (const token of tokenize(base, granularity)) {
    let first = token.start
    for (let k = token.start + 1; k <= token.end; k++) {
      const cut =
        k === token.end || at[k] !== at[k - 1]! + 1 || states[at[k]!] !== states[at[k - 1]!]
      if (!cut) continue
      const start = at[first]!
      const end = at[k - 1]! + 1
      out.push({
        text: base.slice(first, k),
        start,
        end,
        // a piece of a longer word: the word has no equal to be compared with
        fragment: word(start) && (word(start - 1) || word(end)),
      })
      first = k
    }
  }
  return out
}

/**
 * The token differences of two texts that hold tracked changes. Each side is
 * compared as it was before its pending insertions, so a change one side
 * tracks is not a difference; the inserted text of each side then goes in
 * where its own document has it (an insertion of the original as text only
 * the original has, one of the revised document as text only it has).
 * Rejecting every change of the result therefore gives the original as it
 * was before its own.
 */
function trackedTokenOps(
  leftText: string,
  rightText: string,
  granularity: CompareGranularity,
  tracked: TrackedFlags,
  rules: Rules,
): TokenOp[] {
  const words = tokenOps(
    baseTokens(leftText, tracked.left, granularity),
    baseTokens(rightText, tracked.right, granularity),
    rules,
  )
  // a word one side tracks only a piece of has no equal: compare the stretch by character
  const ops: TokenOp[] = []
  for (const stretch of stretches(words)) {
    const gone = stretch.flatMap((op) => (op.kind === 'del' ? [op.l as SideToken] : []))
    const come = stretch.flatMap((op) => (op.kind === 'ins' ? [op.r as SideToken] : []))
    if (gone.length && come.length && [...gone, ...come].some((t) => t.fragment)) {
      ops.push(
        ...tokenOps(
          baseTokens(leftText, tracked.left, 'char', gone[0]!.start, gone[gone.length - 1]!.end),
          baseTokens(rightText, tracked.right, 'char', come[0]!.start, come[come.length - 1]!.end),
          rules,
        ),
      )
    } else ops.push(...stretch)
  }

  // the pending insertions of each side, where its own text has them
  const placed: TokenOp[] = []
  let li = 0
  let ri = 0
  const leftUpTo = (to: number) => {
    if (to > li)
      placed.push({ kind: 'del', l: { text: leftText.slice(li, to), start: li, end: to } })
  }
  const rightUpTo = (to: number) => {
    if (to > ri)
      placed.push({ kind: 'ins', r: { text: rightText.slice(ri, to), start: ri, end: to } })
  }
  for (const op of ops) {
    if (op.kind !== 'ins') {
      leftUpTo(op.l.start)
      li = op.l.end
    }
    if (op.kind !== 'del') {
      rightUpTo(op.r.start)
      ri = op.r.end
    }
    placed.push(op)
  }
  leftUpTo(leftText.length)
  rightUpTo(rightText.length)

  // text one side inserts as a change and the other simply has is the same text, not two
  const out: TokenOp[] = []
  for (const stretch of stretches(placed)) {
    const gone = stretch.flatMap((op) => (op.kind === 'del' ? [op.l] : []))
    const come = stretch.flatMap((op) => (op.kind === 'ins' ? [op.r] : []))
    const whole = (tokens: Token[]) =>
      tokens.every((t, i) => i === 0 || t.start === tokens[i - 1]!.end)
    const pending =
      gone.some((t) => tracked.left[t.start] === 'ins') ||
      come.some((t) => tracked.right[t.start] === 'ins')
    const text = (tokens: Token[]) => tokens.map((t) => t.text).join('')
    if (
      pending &&
      gone.length &&
      come.length &&
      whole(gone) &&
      whole(come) &&
      text(gone) === text(come)
    ) {
      out.push({
        kind: 'eq',
        l: { text: text(gone), start: gone[0]!.start, end: gone[gone.length - 1]!.end },
        r: { text: text(come), start: come[0]!.start, end: come[come.length - 1]!.end },
      })
    } else out.push(...stretch)
  }
  return out
}

/** the ops cut into runs of differences and single equal tokens */
function stretches(ops: TokenOp[]): TokenOp[][] {
  const out: TokenOp[][] = []
  let open: TokenOp[] | null = null
  for (const op of ops) {
    if (op.kind === 'eq') {
      open = null
      out.push([op])
    } else if (open) open.push(op)
    else out.push((open = [op]))
  }
  return out
}

/** neighbouring deletions (or insertions) become one revision region */
function mergeChunks(chunks: Chunk[]): Chunk[] {
  const out: Chunk[] = []
  for (const chunk of chunks) {
    const last = out[out.length - 1]
    if (last && last.kind === 'fmt' && chunk.kind === 'fmt') {
      if (
        last.leftRun === chunk.leftRun &&
        last.left![1] === chunk.left![0] &&
        last.right![1] === chunk.right![0]
      ) {
        last.left![1] = chunk.left![1]
        last.right![1] = chunk.right![1]
        continue
      }
    } else if (last && last.kind === chunk.kind && chunk.kind !== 'keep') {
      const lastRange = chunk.kind === 'del' ? last.left! : last.right!
      const range = chunk.kind === 'del' ? chunk.left! : chunk.right!
      if (lastRange[1] === range[0]) {
        lastRange[1] = range[1]
        continue
      }
    }
    out.push({
      ...chunk,
      ...(chunk.left ? { left: [...chunk.left] as [number, number] } : {}),
      ...(chunk.right ? { right: [...chunk.right] as [number, number] } : {}),
    })
  }
  return out
}

function hasRevision(chunks: Chunk[]): boolean {
  return chunks.some((c) => c.kind !== 'keep' || c.carry)
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

/** the code of the field the token starts in; null outside a field */
function fieldCodeOf(runs: Run[], starts: number[], token: Token): string | null {
  const run = runs[runIndexAt(starts, token.start)]
  const code = run?.instrField ?? run?.refInstr ?? (run?.refField ? `REF ${run.refField}` : null)
  return code === null || code === undefined ? null : code.replace(/\s+/g, ' ').trim()
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
function deletedBlock(
  block: Block,
  author: string,
  date: string,
  ids: RevisionIds,
  carry?: Carrier,
): SaveBlock {
  const revision: RevisionInfo = { author, date, id: ids.take() }
  if (block.runs) {
    // text the original already struck keeps that record; text it inserted is struck on top
    const runs = carry
      ? foreignRuns(markupRuns(block.runs)).map((r) =>
          r.del ? { ...r, ...carry.marks(r) } : { ...r, ...carry.marks(r), del: revision },
        )
      : foreignRuns(comparableRuns(block.runs)).map((r) => ({ ...r, del: revision }))
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
function insertedBlock(
  block: Block,
  author: string,
  date: string,
  ids: RevisionIds,
  keepRevisions = false,
): SaveBlock {
  const revision: RevisionInfo = { author, date, id: ids.take() }
  if (block.runs) {
    const runs = keepRevisions
      ? markupRuns(block.runs).map((r) =>
          isTracked(r) ? emittedRun(r) : { ...emittedRun(r), ins: revision },
        )
      : comparableRuns(block.runs).map((r) => ({ ...emittedRun(r), ins: revision }))
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

/** the runs with their pending changes as they are: deleted text stays, marked deleted */
function markupRuns(runs: Run[]): Run[] {
  return runs.map((run) => (run.image && !run.text ? { ...run, text: PICTURE_TOKEN } : run))
}

/** a run of the revised document, the picture token folded back into the picture run */
function emittedRun(run: Run): Run {
  return run.image && run.text === PICTURE_TOKEN ? { ...run, text: '' } : run
}

/**
 * Runs of the original document that end up in the redline as deleted text.
 * What points into the original package stays behind: a picture is named
 * instead of carried, a note reference goes, and a comment range, a link or a
 * field leaves its text.
 */
function foreignRuns(runs: Run[]): Run[] {
  const out: Run[] = []
  for (const run of runs) {
    if (run.noteRef) continue
    const {
      image,
      commentIds: _commentIds,
      link: _link,
      xeTerm: _xeTerm,
      taInstr: _taInstr,
      taRuns: _taRuns,
      tcInstr: _tcInstr,
      refField: _refField,
      refInstr: _refInstr,
      instrField: _instrField,
      zoteroFieldId: _zoteroFieldId,
      zoteroFieldPart: _zoteroFieldPart,
      fldBeginXml: _fldBeginXml,
      fldDirty: _fldDirty,
      sdtCheckboxXml: _sdtCheckboxXml,
      ...rest
    } = run
    out.push(image ? { ...rest, text: '[picture]' } : rest)
  }
  return out
}

/** drop the drawings, VML pictures and embedded objects of a fragment from the other package */
function withoutForeignObjects(xml: string): string {
  return xml
    .replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, '')
    .replace(/<w:pict(?:\s[^>]*)?>[\s\S]*?<\/w:pict>/g, '')
    .replace(/<w:object(?:\s[^>]*)?>[\s\S]*?<\/w:object>/g, '')
    .replace(/<w:(?:footnote|endnote|comment)Reference\b[^>]*\/>/g, '')
    .replace(/<w:commentRange(?:Start|End)\b[^>]*\/>/g, '')
}

/**
 * Revision ids are unique across the document: they start above every id the
 * revised document already uses, so the redline never repeats one.
 */
class RevisionIds {
  private next: number

  constructor(parts: string[]) {
    let max = 0
    for (const xml of parts) {
      for (const m of xml.matchAll(/\bw:id\s*=\s*["'](\d+)["']/g)) {
        max = Math.max(max, parseInt(m[1]!, 10))
      }
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

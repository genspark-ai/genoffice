import type {
  Block,
  HfParagraph,
  NoteRun,
  ParsedDoc,
  Run,
  TableModel,
  TextboxDisplay,
} from '@genoffice/docx-engine'
import type { RevisionRange } from './revisions'

/**
 * Word's Markup views: Simple Markup (final text + changed-line bars), All
 * Markup (inline revisions), No Markup (as accepted), Original (as rejected).
 */
export type RevisionDisplayMode = 'simple' | 'all' | 'none' | 'original'

/**
 * Word's Show Markup → Balloons: revisions (deletions + format changes +
 * comments in balloons), inline (no balloons; comments as [initials N]
 * markers), comments (Word's default: comments + format changes only).
 */
export type BalloonMode = 'revisions' | 'inline' | 'comments'

export const DEFAULT_BALLOON_MODE: BalloonMode = 'comments'
export const BALLOON_MODE_KEY = 'aidocs.balloonMode'

export function loadBalloonMode(): BalloonMode {
  if (typeof localStorage === 'undefined') return DEFAULT_BALLOON_MODE
  const v = localStorage.getItem(BALLOON_MODE_KEY)
  return v === 'revisions' || v === 'inline' || v === 'comments' ? v : DEFAULT_BALLOON_MODE
}

export type BalloonRevKind = 'del' | 'fmt'

/** which balloon a tracked-change range would get: deletions leave the flow, property changes are 'fmt' */
export function balloonKindOf(kind: string): BalloonRevKind | null {
  if (kind.endsWith('Change') || kind === 'cellMerge') return 'fmt'
  return /^(del|paraMarkDel|moveFrom|rowDel|cellDel|blockDel)$/.test(kind) ? 'del' : null
}

/** tracked-revision balloon kinds the margin shows in this mode */
export function balloonRevKinds(mode: BalloonMode): ReadonlySet<BalloonRevKind> {
  if (mode === 'revisions') return new Set(['del', 'fmt'])
  if (mode === 'comments') return new Set(['fmt'])
  return new Set()
}

/** the balloon mode the app root currently carries (data-balloons; absent outside All Markup print layout) */
export function currentBalloonMode(root: Element | null): BalloonMode | null {
  const v = root?.closest('[data-balloons]')?.getAttribute('data-balloons')
  return v === 'revisions' || v === 'inline' || v === 'comments' ? v : null
}

const HIDDEN_AS_ACCEPTED = new Set<RevisionRange['kind']>([
  'del',
  'moveFrom',
  'rowDel',
  'cellDel',
  'blockDel',
])
const HIDDEN_AS_REJECTED = new Set<RevisionRange['kind']>([
  'ins',
  'moveTo',
  'rowIns',
  'cellIns',
  'blockIns',
])

/** the revision's text leaves the flow in this view (navigation must fall back to All Markup) */
export function revisionHiddenIn(
  mode: RevisionDisplayMode,
  rev: Pick<RevisionRange, 'kind' | 'overlapped'>,
): boolean {
  if (mode === 'all') return false
  if (rev.overlapped) return true
  if (mode === 'original') return HIDDEN_AS_REJECTED.has(rev.kind)
  return HIDDEN_AS_ACCEPTED.has(rev.kind)
}

/** a run outside the main editor (note, header/footer, textbox) leaves the view in this mode */
export function runHiddenIn(mode: RevisionDisplayMode, run: Pick<Run, 'ins' | 'del'>): boolean {
  if (mode === 'all') return false
  if (mode === 'original') return !!run.ins && !run.del
  return !!run.del
}

/** note display runs as the view shows them (the measurement host sits outside the mode root) */
export function noteRunsInView(
  mode: RevisionDisplayMode,
  richParas: NoteRun[][] | undefined,
): NoteRun[][] | undefined {
  if (mode === 'all' || !richParas) return richParas
  return richParas.map((para) => para.filter((r) => !runHiddenIn(mode, r)))
}

/** the Markup view the app root currently carries (rev-display-* class, All when absent) */
export function currentRevisionDisplay(): RevisionDisplayMode {
  if (typeof document === 'undefined') return 'all'
  const root = document.querySelector('.app')
  for (const mode of ['none', 'simple', 'original'] as const) {
    if (root?.classList.contains(`rev-display-${mode}`)) return mode
  }
  return 'all'
}

/**
 * Word prints comment balloons and inline revisions only from the All Markup
 * view (Print Layout); Simple / No Markup / Original print the clean sheet.
 */
export function markupPrintsIn(mode: RevisionDisplayMode, viewMode: string): boolean {
  return mode === 'all' && viewMode === 'print'
}

const runsHaveRevisions = (runs: readonly Run[] | undefined) =>
  !!runs?.some((r) => r.ins || r.del || r.rPrChange)

function textboxesHaveRevisions(boxes: readonly TextboxDisplay[] | undefined): boolean {
  return !!boxes?.some((box) =>
    box.paras.some(
      (p) => runsHaveRevisions(p.runs) || p.cells?.some((c) => c.paras.some(runsHaveRevisions)),
    ),
  )
}

function tableHasRevisions(table: TableModel): boolean {
  if (table.rowRevisions?.some((r) => r)) return true
  if (table.tblPrChange || table.tblGridChange) return true
  if (table.rowPrChanges?.some(Boolean) || table.rowTblPrExChanges?.some(Boolean)) return true
  for (const row of table.rows) {
    for (const cell of row) {
      if (cell.cellRevision || cell.tcPrChange || cell.cellMerge) return true
      for (const p of cell.richParas ?? []) {
        if (runsHaveRevisions(p.runs)) return true
      }
      if (textboxesHaveRevisions(cell.anchoredBoxes)) return true
      if (cell.nestedTables?.some(tableHasRevisions)) return true
    }
  }
  return false
}

/** any w:ins / w:del / w:rPrChange / w:pPrChange (incl. moves, row/cell/block revisions) in the parsed body */
export function blocksHaveRevisions(blocks: Block[]): boolean {
  for (const b of blocks) {
    if (
      b.pPrChangeInfo ||
      b.paraMarkDel ||
      b.paraMarkIns ||
      b.paraMarkRPrChange ||
      b.numberingChange ||
      b.sectPrChange ||
      b.moveRevision ||
      b.blockRevision
    )
      return true
    if (runsHaveRevisions(b.runs)) return true
    if (textboxesHaveRevisions(b.textboxes)) return true
    if (b.table && tableHasRevisions(b.table)) return true
  }
  return false
}

function hfParasHaveRevisions(paras: readonly HfParagraph[] | null | undefined): boolean {
  return !!paras?.some(
    (p) => runsHaveRevisions(p.runs) || p.cells?.some((c) => c.paras.some(runsHaveRevisions)),
  )
}

/** blocksHaveRevisions plus the parts the body does not carry: notes, headers/footers */
export function documentHasRevisions(
  parsed: Pick<
    ParsedDoc,
    | 'blocks'
    | 'footnotes'
    | 'endnotes'
    | 'headerParas'
    | 'footerParas'
    | 'headerFirst'
    | 'footerFirst'
    | 'headerEven'
    | 'footerEven'
    | 'hfParts'
  >,
): boolean {
  if (blocksHaveRevisions(parsed.blocks)) return true
  for (const note of [...parsed.footnotes, ...parsed.endnotes]) {
    if (note.richParas?.some((p) => p.some((r) => r.ins || r.del))) return true
  }
  const parts = [
    parsed.headerParas,
    parsed.footerParas,
    parsed.headerFirst?.paras,
    parsed.footerFirst?.paras,
    parsed.headerEven?.paras,
    parsed.footerEven?.paras,
    ...Object.values(parsed.hfParts ?? {}).map((p) => p.paras),
  ]
  return parts.some(hfParasHaveRevisions)
}

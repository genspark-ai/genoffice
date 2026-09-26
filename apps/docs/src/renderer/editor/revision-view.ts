import type { Block, TableModel } from '@genoffice/docx-engine'
import type { RevisionRange } from './revisions'

/**
 * Word's Markup views: Simple Markup (final text + changed-line bars), All
 * Markup (inline revisions), No Markup (as accepted), Original (as rejected).
 */
export type RevisionDisplayMode = 'simple' | 'all' | 'none' | 'original'

const HIDDEN_AS_ACCEPTED = new Set<RevisionRange['kind']>([
  'del',
  'both',
  'moveFrom',
  'rowDel',
  'cellDel',
  'blockDel',
])
const HIDDEN_AS_REJECTED = new Set<RevisionRange['kind']>([
  'ins',
  'both',
  'moveTo',
  'rowIns',
  'cellIns',
  'blockIns',
])

/** the revision's text leaves the flow in this view (navigation must fall back to All Markup) */
export function revisionHiddenIn(mode: RevisionDisplayMode, kind: RevisionRange['kind']): boolean {
  if (mode === 'all') return false
  if (mode === 'original') return HIDDEN_AS_REJECTED.has(kind)
  return HIDDEN_AS_ACCEPTED.has(kind)
}

/**
 * Word opens a document that carries tracked changes in Simple Markup; a view
 * the user picked explicitly this session wins over that default.
 */
export function initialMarkupMode(
  current: RevisionDisplayMode,
  userPicked: boolean,
  hasRevisions: boolean,
): RevisionDisplayMode {
  if (userPicked || !hasRevisions) return current
  return current === 'all' ? 'simple' : current
}

function tableHasRevisions(table: TableModel): boolean {
  if (table.rowRevisions?.some((r) => r)) return true
  for (const row of table.rows) {
    for (const cell of row) {
      if (cell.cellRevision) return true
      for (const p of cell.richParas ?? []) {
        if (p.runs.some((r) => r.ins || r.del || r.rPrChange)) return true
      }
      if (cell.nestedTables?.some(tableHasRevisions)) return true
    }
  }
  return false
}

/** any w:ins / w:del / w:rPrChange / w:pPrChange (incl. moves, row/cell/block revisions) in the parsed body */
export function blocksHaveRevisions(blocks: Block[]): boolean {
  for (const b of blocks) {
    if (b.pPrChangeInfo || b.paraMarkDel || b.moveRevision || b.blockRevision) return true
    if (b.runs?.some((r) => r.ins || r.del || r.rPrChange)) return true
    if (b.table && tableHasRevisions(b.table)) return true
  }
  return false
}

import type { Block, TableModel } from '@genoffice/docx-engine'

/**
 * Opening tiers by document weight = blocks + characters / CHARS_PER_BLOCK.
 *
 * Measured 2026-09 (Apple silicon, renderer working set after pagination):
 * short synthetic paragraphs 30k blocks / 3 M chars → 1.26 GB, 1945 pages,
 * 17 s; long synthetic paragraphs 10k blocks / 7 M chars → 1.5 GB, 2038
 * pages, 11 s; 20k blocks / 14 M chars → 2.8 GB, 4080 pages, 32 s. That fits
 * about 300 MB + 17 KB per block + 150 B per character, so a block counts as
 * much as ~110 characters and the thresholds are expressed in that unit:
 * 60k short blocks (the earlier read-only line, ~1.1 GB) weigh ~115k.
 * The read-only tier keeps Read Mode's lighter path (the user can leave it);
 * past the refuse line the renderer would go down with the window instead.
 */
export const CHARS_PER_BLOCK = 110
export const LARGE_DOC_READ_ONLY_WEIGHT = 115_000
export const HUGE_DOC_REFUSE_WEIGHT = 230_000
/**
 * Above this the document opens with check-as-you-type spelling off: Chromium
 * re-checks whole paragraphs on every edit, which was 40% of the main-thread
 * time of a paragraph split on a 324-page report (weight ~17k) and nothing
 * measurable on an 81-page one (~4k). Review › Spelling turns it back on.
 */
export const LITE_DOC_WEIGHT = 4_000

export type OpenTier = 'normal' | 'lite' | 'readOnly' | 'refuse'

export function openTierFor(weight: number): OpenTier {
  if (weight > HUGE_DOC_REFUSE_WEIGHT) return 'refuse'
  if (weight > LARGE_DOC_READ_ONLY_WEIGHT) return 'readOnly'
  if (weight > LITE_DOC_WEIGHT) return 'lite'
  return 'normal'
}

export function docWeight(blocks: Block[]): number {
  return blocks.length + Math.round(docTextLength(blocks) / CHARS_PER_BLOCK)
}

/** characters of body text: paragraph runs and table cell text (nested tables included) */
export function docTextLength(blocks: Block[]): number {
  let n = 0
  for (const b of blocks) {
    if (b.runs) for (const r of b.runs) n += r.text.length
    if (b.table) n += tableTextLength(b.table)
  }
  return n
}

function tableTextLength(table: TableModel): number {
  let n = 0
  for (const row of table.rows) {
    for (const cell of row) {
      for (const p of cell.paras) n += p.length
      for (const nested of cell.nestedTables ?? []) n += tableTextLength(nested)
    }
  }
  return n
}

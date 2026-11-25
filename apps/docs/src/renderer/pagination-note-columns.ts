/**
 * Footnote areas laid out in columns. The slicer charges a note with the line
 * holding its reference. An area in columns is as tall as its tallest column, so
 * a note is charged what it adds to its page's area: as the slicer opens a page,
 * every note after it is charged as continuing that page (SliceOutputs.notePage).
 */
import { noteAreaIncrements, type NoteColumns } from './note-columns'
import { noteReserveKey, pageOfY } from './pagination-slices'
import type { BlockBox, NoteBand, PageSlice, SliceOutputs } from './pagination-types'

/** re-slicings a pass may take to reserve the area in every text column of a page */
export const TEXT_COLUMN_ROUNDS = 6

/** what the last pass took for pages of text columns: re-slicings, and whether the reserves held */
export interface NoteColumnsStats {
  rounds: number
  stop: 'settled' | 'cap'
}

interface Site {
  block: BlockBox
  band: NoteBand
}

function setBand(site: Site, height: number): void {
  const d = height - site.band.height
  if (Math.abs(d) < 1e-6) return
  site.band.height = height
  site.block.height += d
  site.block.footnoteExtraPx = (site.block.footnoteExtraPx ?? 0) + d
  const row = site.band.row === undefined ? undefined : site.block.tableRows?.[site.band.row]
  if (row) row.notesPx = (row.notesPx ?? 0) + d
}

const refY = (s: Site) => s.block.top + (s.block.spaceBeforePx ?? 0) + s.band.offset
const sum = (list: readonly number[]) => list.reduce((a, b) => a + b, 0)

/**
 * Sets `out.notePage`. Returns what to run on the slicing: under text columns the
 * area runs below all of them, so each column also gives up what the notes of the
 * other columns add (SliceOutputs.noteReserve); it slices again while a reserve
 * grows. Null when no block carries note lines.
 */
export function noteColumnsPass(
  blocks: readonly BlockBox[],
  out: SliceOutputs,
  colsOf: (section: number) => NoteColumns | null,
  separatorPx: number,
  pageHeightPx: number,
  stats?: NoteColumnsStats,
): ((slices: PageSlice[], reslice: () => PageSlice[]) => PageSlice[]) | null {
  const sites: Site[] = []
  // per block: its own sites [from, to) in `sites`
  const range = new Map<BlockBox, [number, number]>()
  for (const block of blocks) {
    const from = sites.length
    for (const band of block.noteBands ?? []) if (band.lines) sites.push({ block, band })
    range.set(block, [from, sites.length])
  }
  if (sites.length === 0) return null
  const lines = sites.map((s) => s.band.lines!)
  const reserve = new Map<number, number>()
  out.noteReserve = reserve

  out.notePage = (block, y, section, first) => {
    const [from, to] = (block && range.get(block)) ?? [sites.length, sites.length]
    // the notes the page opens with were charged on the page before: their lines carry that
    let open = first ? from : to
    if (!first) while (open > from && refY(sites[open - 1]) >= y - 0.5) open--
    const carried = first ? 0 : to - open
    // the section of the page's first note lays the area out
    const count = colsOf(sites[open]?.block.section ?? section)?.count ?? 1
    const adds = noteAreaIncrements(lines.slice(open), count, pageHeightPx)
    for (let k = carried; k < adds.length; k++) setBand(sites[open + k], adds[k])
    let charged = 0
    for (let k = 0; k < carried; k++) charged += sites[open + k].band.height
    return { shiftPx: sum(adds.slice(0, carried)) - charged, withNotes: carried > 0 }
  }

  return (slices, reslice) => {
    let rounds = 0
    let stop: NoteColumnsStats['stop'] = 'settled'
    for (;;) {
      const pages = out.preParity ?? slices
      const byPage = new Map<number, Site[]>()
      for (const s of sites) {
        const page = pageOfY(pages, refY(s))
        byPage.set(page, [...(byPage.get(page) ?? []), s])
      }
      let grew = false
      for (const [page, list] of byPage) {
        const slice = pages[page]
        const columns = slice?.regions?.length === 1 ? slice.regions[0].columns : undefined
        if (!columns || columns.length < 2) continue
        const colOf = (s: Site) => {
          const at = columns.findIndex((c) => refY(s) + 0.5 >= c.start && refY(s) + 0.5 < c.end)
          return at < 0 ? columns.length - 1 : at
        }
        const firstCol = colOf(list[0])
        columns.forEach((_, c) => {
          const others = sum(list.filter((s) => colOf(s) !== c).map((s) => s.band.height))
          // a column ahead of the first reference is not charged the separator either
          const take = others + (c < firstCol ? separatorPx : 0)
          const key = noteReserveKey(slice.start, c)
          if (take > (reserve.get(key) ?? 0) + 0.5) {
            reserve.set(key, take)
            grew = true
          }
        })
      }
      if (!grew) break
      if (rounds >= TEXT_COLUMN_ROUNDS) {
        stop = 'cap'
        break
      }
      slices = reslice()
      rounds++
    }
    if (stats) Object.assign(stats, { rounds, stop })
    return slices
  }
}

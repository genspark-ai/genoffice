// Whole-document passes run after the block list is built: protected leading
// breaks, balanced DBCS spacing, image z-order normalization.
import type { Block, Run, TableModel, TextboxDisplay } from './types'

/**
 * Word writes wp:anchor relativeHeight as 251658240 + rank, which decodes to
 * small z-orders; other producers write arbitrary values (LibreOffice: 1, 2,
 * …) that decode to huge magnitudes, defeating the editor's ±1 reorder steps
 * and its CSS bands. When any decoded rank is wild, re-rank every anchored
 * image by its decoded value (stable by document order) starting at 0; rank 0
 * is the base level, so its attribute is dropped like an untouched anchor.
 */
/**
 * A protected image block swallows the paragraph's runs, so a page-break run
 * written before the drawing (`<w:p><w:r><w:br w:type="page"/></w:r><w:r>
 * <w:drawing>...`) would silently vanish while Word turns the page there.
 * Nothing visible can sit between such a break and the drawing, so the
 * paragraph-level pageBreakBefore is an equivalent model.
 */
export function applyProtectedLeadingBreaks(blocks: Block[]): void {
  for (const b of blocks) {
    if (b.type !== 'image' || b.format?.pageBreakBefore) continue
    const xml = b.originalXml
    if (!xml) continue
    const drawing = xml.search(/<w:drawing[\s>]|<w:pict[\s>]|<w:object[\s>]/)
    const head = xml.slice(0, drawing === -1 ? xml.length : drawing)
    const br = head.search(/<w:br\s[^>]*w:type="page"/)
    if (br === -1) continue
    const textBefore = [...head.slice(0, br).matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)]
      .map((m) => m[1])
      .join('')
    if (textBefore.trim() === '') b.format = { ...b.format, pageBreakBefore: true }
  }
}

/** wide (double-byte) glyph fraction of a run's text (CJK/hangul/fullwidth forms) */
function wideGlyphFraction(text: string): number {
  let wide = 0
  let n = 0
  for (const ch of text) {
    const cp = ch.codePointAt(0)!
    if (
      (cp >= 0x1100 && cp <= 0x11ff) ||
      (cp >= 0x2e80 && cp <= 0x9fff) ||
      (cp >= 0xac00 && cp <= 0xd7af) ||
      (cp >= 0xf900 && cp <= 0xfaff) ||
      (cp >= 0xff00 && cp <= 0xff60) ||
      cp >= 0x20000
    )
      wide++
    n++
  }
  return n > 0 ? wide / n : 0
}

/**
 * settings.xml balanceSingleByteDoubleByteWidth (standard in HWP-exported
 * docx) makes rPr w:spacing count double on double-byte characters: a Word
 * probe (2026-08-24) moves hangul advances 2pt per 20 twips with the flag and
 * 1pt without, Latin runs unchanged either way. CSS letter-spacing cannot
 * vary per character, so scale the display-only charSpacingTwips by each
 * run's wide-glyph mix (the charScaleEm approximation precedent); saving
 * stays byte-faithful through rawRPr.
 */
export function applyBalancedDbcsSpacing(runGroups: Array<Run[] | undefined>): void {
  for (const runs of runGroups) {
    if (!runs) continue
    for (const r of runs) {
      if (!r.charSpacingTwips || !r.text) continue
      const frac = wideGlyphFraction(r.text)
      if (frac > 0) r.charSpacingTwips = Math.round(r.charSpacingTwips * (1 + frac) * 10) / 10
    }
  }
}

/** every run container reachable from the parsed blocks (tables, textboxes, nested tables) */
export function blockRunGroups(blocks: Block[]): Array<Run[] | undefined> {
  const groups: Array<Run[] | undefined> = []
  const fromBoxes = (boxes?: TextboxDisplay[]) => {
    for (const box of boxes ?? []) for (const para of box.paras) groups.push(para.runs)
  }
  const fromTable = (table?: TableModel) => {
    if (!table) return
    for (const row of table.rows) {
      for (const cell of row) {
        for (const para of cell.richParas ?? []) groups.push(para.runs)
        fromBoxes(cell.anchoredBoxes)
        for (const nested of cell.nestedTables ?? []) fromTable(nested)
      }
    }
  }
  for (const b of blocks) {
    groups.push(b.runs)
    groups.push(b.strayRuns)
    fromTable(b.table ?? undefined)
    fromBoxes(b.textboxes)
  }
  return groups
}

export function normalizeImageZOrders(blocks: Block[]): void {
  // images and floating shapes share Word's z space: rank them together so
  // cross-type overlaps (photo over a background shape) keep their order
  const anchored: Array<{ get: () => number; set: (rank: number) => void }> = []
  for (const b of blocks) {
    if (b.imageZOrder !== undefined) {
      anchored.push({
        get: () => b.imageZOrder!,
        set: (rank) => {
          if (rank === 0) delete b.imageZOrder
          else b.imageZOrder = rank
          // raw XML still carries the wild value; flag for save-time harmonization
          b.imageZOrderNormalized = true
        },
      })
    }
    for (const box of b.textboxes ?? []) {
      if (box.z !== undefined) {
        anchored.push({
          get: () => box.z!,
          // display-only: the box's XML keeps its raw relativeHeight
          set: (rank) => (box.z = rank),
        })
      }
    }
  }
  if (!anchored.some((e) => Math.abs(e.get()) > 10000)) return
  anchored
    .map((e, i) => ({ e, i }))
    .sort((x, y) => x.e.get() - y.e.get() || x.i - y.i)
    .forEach(({ e }, rank) => e.set(rank))
}

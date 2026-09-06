// Whole-document passes run after the block list is built: protected leading
// breaks, balanced DBCS spacing, image z-order normalization.
import type { Block, Run, SectionSettings, TableModel, TextboxDisplay } from './types'
import { EMU_PER_PX } from './parse-xml-text'

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

/** Word probe (2026-09-03): Han and fullwidth forms always take the doubled spacing,
 *  hangul/kana only in legacy faces (Batang, Gulim, Dotum, NanumMyeongjo, Meiryo, Yu,
 *  MS Gothic, and any missing font Word substitutes with Batang) — not in these */
const SINGLE_SPACING_HANGUL_KANA_FONTS = new Set([
  'malgungothic',
  'malgungothicsemilight',
  '맑은고딕',
  'nanumgothic',
  '나눔고딕',
  'applesdgothicneo',
  'mspgothic',
  'ｍｓｐゴシック',
])

function keepsHangulKanaSingle(font: string | undefined): boolean {
  return (
    !!font && SINGLE_SPACING_HANGUL_KANA_FONTS.has(font.toLowerCase().replace(/[\s\u3000]+/g, ''))
  )
}

function isHangulOrKana(cp: number): boolean {
  return (
    (cp >= 0x1100 && cp <= 0x11ff) ||
    (cp >= 0x3040 && cp <= 0x30ff) ||
    (cp >= 0x3130 && cp <= 0x318f) ||
    (cp >= 0x31f0 && cp <= 0x31ff) ||
    (cp >= 0xac00 && cp <= 0xd7af)
  )
}

function doubledGlyphFraction(text: string, hangulKanaSingle: boolean): number {
  let wide = 0
  let n = 0
  for (const ch of text) {
    const cp = ch.codePointAt(0)!
    const hangulKana = isHangulOrKana(cp)
    const isWide =
      hangulKana ||
      (cp >= 0x2e80 && cp <= 0x9fff) ||
      (cp >= 0xf900 && cp <= 0xfaff) ||
      (cp >= 0xff00 && cp <= 0xff60) ||
      cp >= 0x20000
    if (isWide && !(hangulKana && hangulKanaSingle)) wide++
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
 * run's doubled-glyph mix (the charScaleEm approximation precedent); saving
 * stays byte-faithful through rawRPr.
 */
export function applyBalancedDbcsSpacing(
  runGroups: Array<Run[] | undefined>,
  defaultFont?: string,
): void {
  for (const runs of runGroups) {
    if (!runs) continue
    for (const r of runs) {
      if (!r.charSpacingTwips || !r.text) continue
      const frac = doubledGlyphFraction(
        r.text,
        keepsHangulKanaSingle(r.eastAsiaFont ?? defaultFont),
      )
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

/** column span (twips from the column's left edge, wrap gaps included) of a
 *  w:tblpPr table block; it anchors to the block that follows it */
export function floatTableColumnSpan(
  block: Block,
  sect: SectionSettings,
): { leftTwips: number; rightTwips: number } | undefined {
  const model = block.type === 'table' ? block.table : undefined
  const pos = model?.floatPos
  if (!pos || !model?.colWidthsTwips?.length) return undefined
  // keyword X (tblpXSpec) resolves at layout time; only numeric X is known here
  if (/\bw:tblpXSpec=/.test(block.originalXml ?? '')) return undefined
  const width = model.colWidthsTwips.reduce((sum, w) => sum + w, 0)
  const gap = pos.distanceTwips ?? {}
  // omitted w:horzAnchor means page-relative
  const x = pos.horzAnchor === 'page' || !pos.horzAnchor ? pos.xTwips - sect.marginLeft : pos.xTwips
  return { leftTwips: x - (gap.left ?? 0), rightTwips: x + width + (gap.right ?? 0) }
}

const isEmptyParagraph = (b: Block): boolean =>
  b.type === 'paragraph' && (b.runs ?? []).every((r) => r.text.trim() === '' && !r.image)

/**
 * Word never lays a table beside anchored objects: a table following a
 * paragraph whose side-wrapped drawings float with zero flow footprint starts
 * below them. Reserve the drawings' band on that paragraph (empty paragraphs in
 * between sit beside the drawings and keep the adjacency).
 */
export function bandWrappedBoxesBeforeTables(blocks: Block[]): void {
  for (let i = 0; i < blocks.length; i++) {
    const boxes = blocks[i].textboxes
    if (!boxes?.length || !boxes.every((b) => b.floating)) continue
    let j = i + 1
    while (j < blocks.length && isEmptyParagraph(blocks[j])) j++
    if (blocks[j]?.type !== 'table') continue
    for (const b of boxes) {
      if (!b.wrapSides || b.bandBottomPx !== undefined || b.pagePinned || b.pageRelV) continue
      if (b.heightPx === undefined) continue
      const top = Math.round((b.offsetYEmu ?? 0) / EMU_PER_PX)
      if (top + b.heightPx <= 0) continue
      b.bandTopPx = top
      b.bandBottomPx = top + b.heightPx
      b.bandBeside = true
    }
  }
}

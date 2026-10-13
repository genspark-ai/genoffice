import { describe, expect, it } from 'vitest'
import { PDFArray, PDFDict, PDFDocument, PDFName } from 'pdf-lib'
import { mergePdfParts } from '../src/main/pdf-merge'

/**
 * The chunked print merge (docs:save-merged-pdf): pdf-lib's copyPages brings
 * the page content over but not the catalog's named destinations, so every TOC
 * entry a fragment emitted would resolve to nothing in the merged file.
 */

/** named destinations in a PDF -> name -> 1-based page number */
function namedDests(pdf: PDFDocument): Map<string, number> {
  const ctx = pdf.context
  const out = new Map<string, number>()
  const dict = ctx.lookup(pdf.catalog.get(PDFName.of('Dests')))
  if (!(dict instanceof PDFDict)) return out
  for (const [key, value] of dict.entries()) {
    const arr = ctx.lookup(value)
    if (!(arr instanceof PDFArray) || arr.size() === 0) continue
    const target = arr.get(0)
    const idx = pdf.getPages().findIndex((p) => p.ref.toString() === target.toString())
    if (idx >= 0) out.set(String(key).replace(/^\//, ''), idx + 1)
  }
  return out
}

/** link annotations: page number + the name they point at */
function links(pdf: PDFDocument): { page: number; dest: string }[] {
  const ctx = pdf.context
  const rows: { page: number; dest: string }[] = []
  pdf.getPages().forEach((p, i) => {
    const annots = ctx.lookup(p.node.get(PDFName.of('Annots')))
    if (!(annots instanceof PDFArray)) return
    for (const ref of annots.asArray()) {
      const a = ctx.lookup(ref)
      if (!(a instanceof PDFDict)) continue
      if (a.get(PDFName.of('Subtype'))?.toString() !== '/Link') continue
      const dest = a.get(PDFName.of('Dest'))
      rows.push({ page: i + 1, dest: String(dest ?? '').replace(/^\//, '') })
    }
  })
  return rows
}

/** a printed fragment: `pageCount` sheets, a named destination per entry, one link on page 1 */
async function fragment(
  pageCount: number,
  dests: Record<string, number>,
  linkName?: string,
): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const pages = []
  for (let i = 0; i < pageCount; i++) pages.push(doc.addPage([200, 200]))
  const destDict = doc.context.obj({})
  for (const [name, page] of Object.entries(dests)) {
    destDict.set(
      PDFName.of(name),
      doc.context.obj([pages[page - 1].ref, PDFName.of('XYZ'), 0, 27, 0]),
    )
  }
  doc.catalog.set(PDFName.of('Dests'), destDict)
  if (linkName) {
    pages[0].node.set(
      PDFName.of('Annots'),
      doc.context.obj([
        doc.context.obj({
          Type: 'Annot',
          Subtype: 'Link',
          Rect: [10, 10, 60, 24],
          Border: [0, 0, 0],
          Dest: PDFName.of(linkName),
        }),
      ]),
    )
  }
  return doc.save()
}

describe('mergePdfParts', () => {
  it('carries named destinations across the merge and repoints them at the copied pages', async () => {
    // two print chunks: the link sits in the first, its target page in the second
    const a = await fragment(2, { chap01: 1, chap02: 2 }, 'chap02')
    const b = await fragment(2, { chap03: 1 }, 'chap03')

    const merged = await PDFDocument.load(await mergePdfParts([a, b]))

    expect(merged.getPageCount()).toBe(4)
    // page numbers are global after the merge, not per fragment
    expect([...namedDests(merged)]).toEqual([
      ['chap01', 1],
      ['chap02', 2],
      ['chap03', 3],
    ])
    // the annotations reference destinations by name and survive the copy
    expect(links(merged)).toEqual([
      { page: 1, dest: 'chap02' },
      { page: 3, dest: 'chap03' },
    ])
  })

  it('keeps the destination view (x/y/zoom), not just the page', async () => {
    const part = await fragment(1, { chap01: 1 })
    const merged = await PDFDocument.load(await mergePdfParts([part]))
    const dict = merged.context.lookup(merged.catalog.get(PDFName.of('Dests')))
    if (!(dict instanceof PDFDict)) throw new Error('merged PDF has no /Dests')
    const arr = merged.context.lookup(dict.get(PDFName.of('chap01')))
    if (!(arr instanceof PDFArray)) throw new Error('chap01 has no destination array')
    expect(String(arr.get(1))).toBe('/XYZ')
    expect(arr.get(3).toString()).toBe('27')
  })

  it('merges fragments that carry no destinations at all', async () => {
    const bare = await PDFDocument.create()
    bare.addPage([200, 200])
    const part = await fragment(1, { chap01: 1 })
    const merged = await PDFDocument.load(await mergePdfParts([await bare.save(), part]))
    expect(merged.getPageCount()).toBe(2)
    expect([...namedDests(merged)]).toEqual([['chap01', 2]])
  })
})

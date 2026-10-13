import { PDFArray, PDFDict, PDFDocument, PDFName, type PDFPage } from 'pdf-lib'

/**
 * Merge printed PDF fragments into one file, in page order.
 *
 * pdf-lib's copyPages brings the page content and annotations across but not the
 * catalog's named destinations, so every fragment's /Dests is dropped and the
 * merged TOC/cross-reference links resolve to nothing. Each fragment's
 * destinations are read before the copy and re-emitted against the copied pages.
 *
 * Annotations point at destinations *by name*, so one merged /Dests also
 * repairs the cross-chunk case for free: a link printed in one chunk carries a
 * name whose target page was display:none in that chunk, and it resolves
 * through the merged map to the page of whichever chunk printed that target.
 */
export async function mergePdfParts(parts: Uint8Array[]): Promise<Uint8Array> {
  const merged = await PDFDocument.create()
  for (const part of parts) {
    const source = await PDFDocument.load(part)
    const destinations = namedDestinations(source)
    const offset = merged.getPageCount()
    const pages = await merged.copyPages(source, source.getPageIndices())
    for (const page of pages) merged.addPage(page)
    for (const [name, dest] of destinations) {
      const page = merged.getPage(offset + dest.pageIndex)
      if (!page) continue
      addDestination(merged, name, page, dest.view)
    }
  }
  return merged.save()
}

interface FragmentDestination {
  /** 0-based page index within its own fragment */
  pageIndex: number
  /** the destination's view arguments (e.g. /XYZ x y zoom), kept verbatim */
  view: PDFArray
}

/** The named destinations a fragment carries, from either catalog spelling. */
function namedDestinations(pdf: PDFDocument): Map<string, FragmentDestination> {
  const ctx = pdf.context
  const out = new Map<string, FragmentDestination>()
  const pages = pdf.getPages()
  const read = (dict: PDFDict | undefined): void => {
    if (!dict) return
    for (const [key, value] of dict.entries()) {
      const arr = ctx.lookup(value)
      if (!(arr instanceof PDFArray) || arr.size() === 0) continue
      const target = arr.get(0)
      const pageIndex = pages.findIndex((p) => p.ref.toString() === target.toString())
      if (pageIndex < 0) continue
      // everything after the page ref is the destination's view (/XYZ …)
      const view = PDFArray.withContext(ctx)
      for (let i = 1; i < arr.size(); i++) view.push(arr.get(i))
      out.set(String(key).replace(/^\//, ''), { pageIndex, view })
    }
  }
  const catalogDests = ctx.lookup(pdf.catalog.get(PDFName.of('Dests')))
  read(catalogDests instanceof PDFDict ? catalogDests : undefined)
  // producers may instead put the name tree under /Names
  const names = ctx.lookup(pdf.catalog.get(PDFName.of('Names')))
  if (names instanceof PDFDict) {
    const tree = ctx.lookup(names.get(PDFName.of('Dests')))
    read(tree instanceof PDFDict ? tree : undefined)
  }
  return out
}

function addDestination(pdf: PDFDocument, name: string, page: PDFPage, view: PDFArray): void {
  const ctx = pdf.context
  const arr = PDFArray.withContext(ctx)
  arr.push(page.ref)
  for (let i = 0; i < view.size(); i++) arr.push(view.get(i))
  const existing = ctx.lookup(pdf.catalog.get(PDFName.of('Dests')))
  const dict = existing instanceof PDFDict ? existing : ctx.obj({})
  if (!(existing instanceof PDFDict)) pdf.catalog.set(PDFName.of('Dests'), dict)
  dict.set(PDFName.of(name), arr)
}

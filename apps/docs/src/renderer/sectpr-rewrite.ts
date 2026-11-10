import type { SaveBlock } from '@genoffice/docx-engine'

export interface SectPrRewrite {
  /** the parsed sectPr fragment as it sits in the block's original bytes */
  from: string
  to: string
  originalXml: string
}

/**
 * Write edited section settings into the save block that carries each
 * section-break paragraph. An untouched paragraph is emitted from its original
 * bytes; one whose text was edited is regenerated and keeps the old sectPr in
 * rawPPr; an editor-built fragment keeps it in its xml. All three must pick up
 * the new sectPr or the file silently keeps the old page setup.
 */
export function applySectPrRewrites(
  saveBlocks: SaveBlock[],
  indexByDocx: Map<number, number>,
  rewrites: Map<number, SectPrRewrite>,
): SaveBlock[] {
  if (rewrites.size === 0) return saveBlocks
  const docxAt = new Map<number, number>()
  for (const [docxIndex, at] of indexByDocx) docxAt.set(at, docxIndex)
  return saveBlocks.map((fb, at) => {
    const docxIndex =
      fb.kind === 'original'
        ? fb.docxIndex
        : fb.kind === 'xml'
          ? (fb.docxIndex ?? docxAt.get(at))
          : docxAt.get(at)
    if (docxIndex === undefined) return fb
    const rw = rewrites.get(docxIndex)
    if (!rw) return fb
    if (fb.kind === 'original') {
      return {
        kind: 'xml',
        xml: rw.originalXml.replace(rw.from, rw.to),
        docxIndex,
        ...(fb.revision ? { revision: fb.revision } : {}),
      }
    }
    if (fb.kind === 'generated' && fb.block.rawPPr?.includes(rw.from)) {
      return { ...fb, block: { ...fb.block, rawPPr: fb.block.rawPPr.replace(rw.from, rw.to) } }
    }
    if (fb.kind === 'xml' && fb.xml.includes(rw.from)) {
      return { ...fb, xml: fb.xml.replace(rw.from, rw.to) }
    }
    return fb
  })
}

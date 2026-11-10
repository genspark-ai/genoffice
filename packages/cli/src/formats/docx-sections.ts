import {
  applySectionStartType,
  sectionFromSectPr,
  type SaveBlock,
  type SaveOptions,
  type SectionInfo,
} from '@genoffice/docx-engine'
import type { OpenDocument } from './docx'

import type {
  AiPageSetupAccess,
  AiSectionState,
} from '../../../../apps/docs/src/renderer/ai/page-setup'
import type { SavePlan } from '../../../../apps/docs/src/renderer/editor/convert'

const SECT_PR = /<w:sectPr[^>]*\/>|<w:sectPr[\s\S]*?<\/w:sectPr>/

/**
 * A section ends at a sectPr: inside an original section-break paragraph, inside
 * a section-break paragraph this batch inserted (genXml), or the trailing hidden
 * one. Edits to original / trailing sectPr wait in `side.sectPr` (by docxIndex)
 * until save; a generated paragraph is patched in place.
 */
interface Boundary {
  kind: 'original' | 'generated' | 'trailing'
  /** top-level PM index of the block carrying the sectPr (childCount for the trailing one) */
  pmIndex: number
  docxIndex?: number
  xml: string
}

function boundaries(doc: OpenDocument): Boundary[] {
  const { parsed, side } = doc
  const byDocx = new Map<number, string>()
  for (const b of parsed.blocks) {
    if (b.docxIndex !== null && b.originalXml?.includes('<w:sectPr')) {
      const m = SECT_PR.exec(b.originalXml)
      if (m) byDocx.set(b.docxIndex, m[0])
    }
  }
  const out: Boundary[] = []
  const pm = doc.editor.state.doc
  pm.forEach((node, _offset, index) => {
    const genXml = node.attrs.genXml
    if (node.type.name === 'docProtected' && typeof genXml === 'string') {
      const m = SECT_PR.exec(genXml)
      if (m) out.push({ kind: 'generated', pmIndex: index, xml: m[0] })
      return
    }
    const di = node.attrs.docxIndex
    if (typeof di === 'number' && byDocx.has(di)) {
      out.push({
        kind: 'original',
        pmIndex: index,
        docxIndex: di,
        xml: side.sectPr.get(di) ?? byDocx.get(di)!,
      })
    }
  })
  const trailing = parsed.blocks.find((b) => b.hidden && b.originalXml?.includes('<w:sectPr'))
  if (trailing && trailing.docxIndex !== null) {
    const original = SECT_PR.exec(trailing.originalXml!)?.[0] ?? ''
    out.push({
      kind: 'trailing',
      pmIndex: pm.childCount,
      docxIndex: trailing.docxIndex,
      xml: side.sectPr.get(trailing.docxIndex) ?? original,
    })
  }
  return out
}

/** SectionInfo per section with firstBlockIndex / lastBlockIndex as live PM indexes */
function sections(doc: OpenDocument): { info: SectionInfo; boundary: Boundary }[] {
  const count = doc.editor.state.doc.childCount
  const out: { info: SectionInfo; boundary: Boundary }[] = []
  let first = 0
  for (const b of boundaries(doc)) {
    const last = b.kind === 'trailing' ? Math.max(count - 1, first) : b.pmIndex
    out.push({ info: sectionFromSectPr(b.xml, first, last, doc.parsed.gutterAtTop), boundary: b })
    first = last + 1
  }
  return out
}

export function listSections(doc: OpenDocument): AiSectionState[] {
  return sections(doc).map(({ info }, i) =>
    doc.mods.pageSetup.describeSection(info, i, info.firstBlockIndex, info.lastBlockIndex),
  )
}

function store(doc: OpenDocument, b: Boundary, xml: string): void {
  if (b.kind === 'generated') {
    const pm = doc.editor.state.doc
    const node = pm.child(b.pmIndex)
    let pos = 0
    for (let i = 0; i < b.pmIndex; i++) pos += pm.child(i).nodeSize
    const genXml = String(node.attrs.genXml).replace(b.xml, xml)
    doc.editor.view.dispatch(
      doc.editor.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, genXml }),
    )
    doc.mods.tools.markDocSeen(doc.editor)
    return
  }
  doc.side.sectPr.set(b.docxIndex!, xml)
}

export function pageSetupAccess(doc: OpenDocument): AiPageSetupAccess {
  const { pageSetup } = doc.mods
  return {
    list: () => listSections(doc),
    current: (index) => sections(doc)[index]?.info,
    set: (index, resolved) => {
      const sec = sections(doc)[index]
      if (!sec) return `section ${index} does not exist`
      store(doc, sec.boundary, pageSetup.applyResolvedPageSetup(sec.boundary.xml, resolved))
      // the header/footer tool reads and may re-write the flag; keep both stores agreeing
      if (sec.boundary.kind === 'trailing' && resolved.titlePg !== undefined) {
        doc.side.titlePg = resolved.titlePg
        doc.side.titlePgDirty = true
      }
      return null
    },
    insertBreak: (type, afterBlockIndex) => {
      const all = sections(doc)
      if (all.length === 0) return 'the document has no section properties to copy'
      const owner =
        all.find((s) => Math.max(afterBlockIndex, 0) <= s.info.lastBlockIndex) ??
        all[all.length - 1]!
      const pm = doc.editor.state.doc
      let pos = 0
      for (let i = 0; i <= afterBlockIndex && i < pm.childCount; i++) pos += pm.child(i).nodeSize
      doc.editor
        .chain()
        .insertContentAt(pos, {
          type: 'docProtected',
          attrs: {
            docxIndex: null,
            blockType: 'passthrough',
            label: 'Section break paragraph',
            previewText: '',
            genXml: pageSetup.sectionBreakParagraphXml(owner.boundary.xml),
          },
        })
        .run()
      // the inserted paragraph shifted the owner's boundary by one block
      const shifted =
        owner.boundary.kind === 'generated' && owner.boundary.pmIndex > afterBlockIndex
          ? { ...owner.boundary, pmIndex: owner.boundary.pmIndex + 1 }
          : owner.boundary
      store(doc, shifted, applySectionStartType(owner.boundary.xml, type))
      doc.mods.tools.markDocSeen(doc.editor)
      return null
    },
  }
}

/**
 * Fold pending sectPr edits into the save: the trailing one travels verbatim as
 * SaveOptions.trailingSectPr (the engine owns the hidden block), the others
 * replace the sectPr inside their section-break paragraph wherever the plan
 * put it: untouched (original bytes), regenerated after a body edit (the
 * paragraph's pPr passthrough) or rewritten as an XML fragment.
 */
export function applySectionEdits(
  doc: OpenDocument,
  plan: Pick<SavePlan, 'saveBlocks' | 'saveBlockIndexByDocx'>,
): { saveBlocks: SaveBlock[]; options: Partial<SaveOptions> } {
  const { parsed, side } = doc
  if (side.sectPr.size === 0) return { saveBlocks: plan.saveBlocks, options: {} }
  const trailing = parsed.blocks.find((b) => b.hidden && b.originalXml?.includes('<w:sectPr'))
  const options: Partial<SaveOptions> = {}
  const out = [...plan.saveBlocks]
  for (const [docxIndex, xml] of side.sectPr) {
    if (trailing && docxIndex === trailing.docxIndex) {
      options.trailingSectPr = xml
      continue
    }
    const at = plan.saveBlockIndexByDocx.get(docxIndex)
    const fb = at === undefined ? undefined : out[at]
    if (!fb) continue
    if (fb.kind === 'original') {
      const block = parsed.blocks.find((b) => b.docxIndex === docxIndex)
      if (block?.originalXml)
        out[at!] = { ...fb, kind: 'xml', xml: block.originalXml.replace(SECT_PR, xml), docxIndex }
    } else if (fb.kind === 'generated' && fb.block.rawPPr && SECT_PR.test(fb.block.rawPPr)) {
      out[at!] = { ...fb, block: { ...fb.block, rawPPr: fb.block.rawPPr.replace(SECT_PR, xml) } }
    } else if (fb.kind === 'xml' && SECT_PR.test(fb.xml)) {
      out[at!] = { ...fb, xml: fb.xml.replace(SECT_PR, xml) }
    }
  }
  return { saveBlocks: out, options }
}

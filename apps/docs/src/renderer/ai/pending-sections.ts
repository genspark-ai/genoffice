import type { Editor } from '@tiptap/core'
import type { Node as PmDocNode } from '@tiptap/pm/model'
import type { SectionInfo } from '@genoffice/docx-engine'
import { markDocSeen } from './tools'

/**
 * Section owning a top-level block. Blocks are mapped by the docxIndex of the
 * nearest original block at or before them; a section-break paragraph inserted
 * this session (`pendingBreak`, its sectPr in a generated paragraph's genXml)
 * ends its section, so blocks after it belong to the following one.
 */
export function sectionIndexAtBlock(
  pmDoc: PmDocNode,
  sections: SectionInfo[],
  topIndex: number,
): number {
  if (sections.length === 0) return 0
  let docxIndex: number | null = null
  for (let i = Math.min(topIndex, pmDoc.childCount - 1); i >= 0; i--) {
    const node = pmDoc.child(i)
    const di = node.attrs?.docxIndex as number | null | undefined
    if (di !== null && di !== undefined) {
      docxIndex = di
      break
    }
    const genXml = node.attrs?.genXml
    if (i < topIndex && typeof genXml === 'string' && genXml.includes('<w:sectPr')) {
      const pending = sectionIndexAtBlock(pmDoc, sections, i)
      if (sections[pending]?.pendingBreak && genXml.includes(sections[pending]!.sectPrXml)) {
        return Math.min(pending + 1, sections.length - 1)
      }
    }
  }
  const s = docxIndex === null ? 0 : sections.findIndex((sec) => docxIndex! <= sec.lastBlockIndex)
  return s >= 0 ? s : sections.length - 1
}

/**
 * Rewrite the sectPr of an unsaved section break in place (the generated
 * paragraph's genXml). Copies of one sectPr can sit in several break
 * paragraphs, so the paragraph is matched by the section it closes, not by
 * its XML alone. False when the paragraph is gone.
 */
export function patchPendingSectPr(
  editor: Editor,
  sections: SectionInfo[],
  index: number,
  nextXml: string,
): boolean {
  const sec = sections[index]
  if (!sec?.pendingBreak) return false
  const pmDoc = editor.state.doc
  let pos = 0
  for (let i = 0; i < pmDoc.childCount; i++) {
    const node = pmDoc.child(i)
    const genXml = node.attrs.genXml
    if (
      node.type.name === 'docProtected' &&
      typeof genXml === 'string' &&
      genXml.includes(sec.sectPrXml) &&
      sectionIndexAtBlock(pmDoc, sections, i) === index
    ) {
      editor.view.dispatch(
        editor.state.tr.setNodeMarkup(pos, undefined, {
          ...node.attrs,
          genXml: genXml.replace(sec.sectPrXml, nextXml),
        }),
      )
      markDocSeen(editor)
      return true
    }
    pos += node.nodeSize
  }
  return false
}

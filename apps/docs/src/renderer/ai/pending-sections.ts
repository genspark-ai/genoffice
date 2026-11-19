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
  if (!rewritePendingGenXml(editor, sections, index, nextXml)) return false
  markDocSeen(editor)
  return true
}

function rewritePendingGenXml(
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
      return true
    }
    pos += node.nodeSize
  }
  return false
}

/**
 * The section a break inserted after top-level block `afterBlock` closes: it
 * owns the owner's blocks up to the break (by docxIndex), the owner keeps the
 * rest. Modelled until a save re-reads the file.
 */
export function pendingSectionAt(
  pmDoc: PmDocNode,
  owner: SectionInfo,
  afterBlock: number,
  sectPrXml: string,
): SectionInfo {
  let lastDocx: number | null = null
  for (let i = Math.min(afterBlock, pmDoc.childCount - 1); i >= 0; i--) {
    const di = pmDoc.child(i).attrs?.docxIndex as number | null | undefined
    if (di !== null && di !== undefined) {
      lastDocx = di
      break
    }
  }
  return {
    ...owner,
    pendingBreak: true,
    sectPrXml,
    lastBlockIndex: lastDocx ?? owner.firstBlockIndex - 1,
  }
}

/**
 * A Layout-tab edit on sections behind unsaved breaks: the save's dirty list
 * rewrites parsed break paragraphs only, so these are written into the
 * generated paragraph now. Returns the list with the sectPr each now holds.
 */
export function rewritePendingSectPrs(
  editor: Editor,
  sections: SectionInfo[],
  xmlOf: (sec: SectionInfo, index: number) => string | null,
): SectionInfo[] {
  const next = [...sections]
  sections.forEach((sec, i) => {
    if (!sec.pendingBreak) return
    const xml = xmlOf(sec, i)
    if (xml && xml !== sec.sectPrXml && rewritePendingGenXml(editor, next, i, xml))
      next[i] = { ...sec, sectPrXml: xml }
  })
  return next
}

/**
 * The model without the unsaved sections whose break paragraph left the
 * document (undo, deletion); the next section takes their blocks back. Breaks
 * are counted by position, not by sectPr text: the model can trail a patch.
 * Null when every modelled break is still there.
 */
export function dropRemovedPendingSections(
  pmDoc: PmDocNode,
  sections: SectionInfo[],
): { sections: SectionInfo[]; dropped: number[] } | null {
  if (!sections.some((s) => s.pendingBreak)) return null
  const breaks: Array<{ after: number; before: number }> = []
  let lastDocx = -1
  pmDoc.forEach((node) => {
    const di = node.attrs?.docxIndex as number | null | undefined
    if (di !== null && di !== undefined) {
      for (const b of breaks) if (b.before === Infinity) b.before = di
      lastDocx = di
      return
    }
    const genXml = node.attrs?.genXml
    if (typeof genXml === 'string' && genXml.includes('<w:sectPr'))
      breaks.push({ after: lastDocx, before: Infinity })
  })
  const out: SectionInfo[] = []
  const dropped: number[] = []
  const kept = new Map<number, number>()
  let carry: Pick<SectionInfo, 'firstBlockIndex' | 'startType'> | null = null
  sections.forEach((s, i) => {
    if (s.pendingBreak) {
      const end = s.lastBlockIndex
      const used = kept.get(end) ?? 0
      if (breaks.filter((b) => b.after <= end && b.before > end).length <= used) {
        dropped.push(i)
        carry ??= { firstBlockIndex: s.firstBlockIndex, startType: s.startType }
        return
      }
      kept.set(end, used + 1)
    }
    out.push(carry ? { ...s, ...carry } : s)
    carry = null
  })
  return dropped.length > 0 ? { sections: out, dropped } : null
}

import type { SectionInfo, TableModel } from '@genoffice/docx-engine'

/** twips a body table may use: growth bound, text column and paper (resolveTableWidth input) */
export interface TableWidthBudget {
  avail: number
  fit: number
  paper: number
}

export function sectionWidthBudget(
  sections: SectionInfo[] | undefined,
  docxIndex: number | null,
): TableWidthBudget | null {
  if (!sections?.length || docxIndex == null) return null
  const sec =
    sections.find((s) => docxIndex >= s.firstBlockIndex && docxIndex <= s.lastBlockIndex) ??
    sections[sections.length - 1]
  const paper = sec.settings.pageWidth
  const avail = paper - sec.settings.marginLeft
  if (avail <= 0) return null
  const fit = avail - sec.settings.marginRight
  return { avail, fit: fit > 0 ? fit : avail, paper }
}

/** centered tables (w:jc or w:tblpXSpec) spill both margins, so their growth bound is the paper width */
export function budgetAvail(model: TableModel, budget: TableWidthBudget): number {
  return model.align === 'center' || model.floatPos?.xSpec === 'center'
    ? budget.paper
    : budget.avail
}

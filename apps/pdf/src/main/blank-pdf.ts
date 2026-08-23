import { PDFDocument } from 'pdf-lib'

/** A4 portrait in PDF points */
const A4_SIZE: [number, number] = [595.28, 841.89]

/**
 * Minimal single-page blank PDF backing the shell's "New PDF" entry — the PDF
 * module has no in-memory blank mode, so the file must exist before it opens
 * (same pattern as the blank workbook in sheets).
 */
export async function blankPdfBuffer(): Promise<Buffer> {
  const doc = await PDFDocument.create()
  doc.addPage(A4_SIZE)
  return Buffer.from(await doc.save())
}

/**
 * The headless paths for the formats beyond docx, for a host that needs to read a
 * workbook, deck or pdf and validate or apply operations without an editor — the
 * shell's script kind dispatches on the document's extension and runs each format
 * through the same engine its command line uses, so a dry run means the same thing
 * in both places.
 *
 * Reads stay bounded at the call sites; the sheets write is the workbook DSL (the
 * in-app propose_operations vocabulary) over the gateway, the deck write is the
 * slides app's own op vocabulary, and the pdf side is read-only.
 */
export { runWorkbookDsl } from './formats/xlsx-dsl'
export type { DslOutcome } from './formats/xlsx-dsl'
export { blankWorkbook, cellEditsFromTable, writeWorkbook } from './formats/xlsx'
export type { WriteOutcome } from './formats/xlsx'
export { applyOps, blankDeck, describeDeck, openDeck, saveDeck } from './formats/pptx'
export type { DeckSummary, ElementSummary, SlideSummary } from './formats/pptx'
export type { Op, TxnResult } from '@genoffice/pptx-ops'
export { auditDeckBytes, deckAuditDetail } from './formats/slide-spec'
export { pdfText, PDF_PAGE_PREVIEW_CHARS } from './formats/pdf'
export type { PdfPageText, PdfText } from './formats/pdf'

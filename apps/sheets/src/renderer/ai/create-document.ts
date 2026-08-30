/// AI create_document, renderer side: xlsx/csv serialize one worksheet's
/// display grid (values as shown, like the manual CSV export) and hand the
/// text to the main process to write into the default save folder; docx/pdf/md
/// forward AI-authored content untouched (the shell routes them into the
/// docs-owned creation flow).

import { csvSheetById, serializeActiveSheetCsv, sheetHasFormulas } from '../csv-export'
import type { LazyWorkbookState, UniverRuntime } from '../univer-state'
import type { CreateDocumentToolOutcome, CreateDocumentToolRequest } from './tools'

export interface AiCreateDocumentContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
}

export async function createAiDocument(
  ctx: AiCreateDocumentContext,
  request: CreateDocumentToolRequest,
): Promise<CreateDocumentToolOutcome> {
  if (request.type === 'docx' || request.type === 'pdf' || request.type === 'md') {
    const result = await window.desktopApi.createDocument(request)
    if (!result.ok) return { ok: false, error: result.error ?? 'creating the document failed' }
    return {
      ok: true,
      name: `${request.title}.${request.type}`,
      ...(result.path ? { path: result.path } : {}),
    }
  }
  const sheet = csvSheetById(ctx.univerRef.current, request.sheetId)
  if (!sheet) {
    return {
      ok: false,
      error: 'The worksheet is not available (no workbook open or unknown sheetId).',
    }
  }
  const state = ctx.lazyWorkbookRef.current
  // File-backed sheets of a not-fully-loaded workbook would serialize only
  // the streamed-in region — refuse instead of silently exporting partial
  // data. Sheets added this session live entirely in the grid and are exempt
  // (same rule as ensureRangeLoaded).
  if (
    state &&
    !state.flags.preloadComplete &&
    state.file.sheets.some((fileSheet) => fileSheet.id === sheet.getSheetId())
  ) {
    return {
      ok: false,
      error:
        "The workbook has not finished loading, so this sheet's data is not fully in the grid yet — retry shortly. " +
        'On very large streamed workbooks, copy the needed rows into a new sheet (add_sheet + copy_range) and export that sheet instead.',
    }
  }
  const content = serializeActiveSheetCsv(sheet, state)
  if (content === 'too-large') {
    return { ok: false, error: 'The worksheet is too large to export as a standalone file.' }
  }
  const sheetName = sheet.getSheetName()
  const title = request.title?.trim() || sheetName
  const result = await window.desktopApi.createDocument({
    type: request.type,
    title,
    content,
    ...(request.type === 'xlsx' ? { sheetName } : {}),
  })
  if (!result.ok) return { ok: false, error: result.error ?? 'creating the file failed' }
  return {
    ok: true,
    name: `${title}.${request.type}`,
    sheetName,
    hadFormulas: sheetHasFormulas(sheet),
    ...(result.path ? { path: result.path } : {}),
  }
}

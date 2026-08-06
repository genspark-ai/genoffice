/// CSV export: writes the renderer-serialized CSV where the save dialog points.

import { writeFile } from 'node:fs/promises'

import { BrowserWindow, dialog } from 'electron'

import type { IpcMainInvokeEvent } from 'electron'
import type {
  WorkbookExportCsvRequest,
  WorkbookExportCsvResult,
} from '../shared/desktop-api'

export async function exportCsv(
  event: IpcMainInvokeEvent,
  request: WorkbookExportCsvRequest,
): Promise<WorkbookExportCsvResult> {
  const parent = BrowserWindow.fromWebContents(event.sender)
  const dialogOptions = {
    defaultPath: request.fileName,
    filters: [{ name: 'CSV', extensions: ['csv'] }],
  }
  const selection = parent
    ? await dialog.showSaveDialog(parent, dialogOptions)
    : await dialog.showSaveDialog(dialogOptions)
  if (selection.canceled || !selection.filePath) return { canceled: true }

  await writeFile(selection.filePath, request.csv, 'utf8')
  return { canceled: false, path: selection.filePath }
}

export type SaveAsExportFormat = 'pdf' | 'html' | 'md'

const EXPORT_FORMAT_BY_EXTENSION: Record<string, SaveAsExportFormat> = {
  pdf: 'pdf',
  html: 'html',
  htm: 'html',
  md: 'md',
}

export const SAVE_AS_EXPORT_FILTERS = [
  { name: 'PDF', extensions: ['pdf'] },
  { name: 'HTML', extensions: ['html'] },
  { name: 'Markdown', extensions: ['md'] },
]

/** The export a Save As destination asks for; null means the document itself is saved there. */
export function saveAsExportFormat(filePath: string): SaveAsExportFormat | null {
  const extension = filePath.toLowerCase().match(/\.([a-z0-9]+)$/)?.[1] ?? ''
  return EXPORT_FORMAT_BY_EXTENSION[extension] ?? null
}

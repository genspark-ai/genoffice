import type { DocumentTabKind } from '../shared/tabs-api'

// shared by the open route, the open-dialog filter and the rename gate so they cannot drift
const DOCX_RE = /\.docx$/i
const XLSX_RE = /\.(xlsx|xlsm|xls|csv|tsv)$/i
const PPTX_RE = /\.pptx$/i
const PDF_RE = /\.pdf$/i
const HTML_RE = /\.html?$/i

const TEXT_RE = /\.(md|markdown|txt|json)$/i

export { DOCX_RE, XLSX_RE, PPTX_RE, PDF_RE, HTML_RE, TEXT_RE }

export function appForPath(path: string): DocumentTabKind | undefined {
  if (DOCX_RE.test(path)) return 'docs'
  if (XLSX_RE.test(path)) return 'sheets'
  if (PPTX_RE.test(path)) return 'slides'
  if (PDF_RE.test(path)) return 'pdf'
  if (TEXT_RE.test(path)) return 'markdown'
  if (HTML_RE.test(path)) return 'html'
  return undefined
}

/** A legal name in an extension no app routes to would turn an openable file into an unopenable one. */
export function renameStaysInApp(from: string, newName: string): boolean {
  const current = appForPath(from)
  if (!current) return true
  return appForPath(newName) === current
}

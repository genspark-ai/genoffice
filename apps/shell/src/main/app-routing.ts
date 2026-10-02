/** Which app opens a path, by file extension. Shared so the open route and the
 *  rename gate cannot drift: a rename that moved a file out of its app would
 *  leave a name the shell can no longer open. */
const DOCX_RE = /\.docx$/i
const XLSX_RE = /\.(xlsx|xlsm|xls|csv|tsv)$/i
const PPTX_RE = /\.pptx$/i
const PDF_RE = /\.pdf$/i
const MD_RE = /\.(md|markdown)$/i
const HTML_RE = /\.html?$/i

/** The text app: Markdown, plus the plain-text and JSON it also opens. */
const TEXT_RE = /\.(md|markdown|txt|json)$/i

export { DOCX_RE, XLSX_RE, PPTX_RE, PDF_RE, MD_RE, HTML_RE, TEXT_RE }

/** The app kind a path belongs to, or undefined when nothing opens it. */
export function appForPath(
  path: string,
): 'docs' | 'sheets' | 'slides' | 'pdf' | 'text' | 'html' | undefined {
  if (DOCX_RE.test(path)) return 'docs'
  if (XLSX_RE.test(path)) return 'sheets'
  if (PPTX_RE.test(path)) return 'slides'
  if (PDF_RE.test(path)) return 'pdf'
  if (TEXT_RE.test(path)) return 'text'
  if (HTML_RE.test(path)) return 'html'
  return undefined
}

/**
 * True when renaming `from` to the base name `newName` leaves the file in the
 * app that already opened it.
 *
 * The character-level rules in rename-validation still apply; this only catches
 * the one rename they cannot see — a perfectly legal name in an extension no
 * app routes to, which turns an openable file into an unopenable one. It also
 * lets note.md become note.markdown, which stays in the text app.
 */
export function renameStaysInApp(from: string, newName: string): boolean {
  const current = appForPath(from)
  // a file no app opened has no app to stay in
  if (!current) return true
  return appForPath(newName) === current
}

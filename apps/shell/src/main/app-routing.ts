import { appForPath, documentApps } from './app-registry'
import type { DocumentTabKind } from '../shared/tabs-api'

/** the routing pattern for one app's extensions, built from the registry's
 *  extension list so the regexes and the table cannot drift */
const extensionRe = (id: DocumentTabKind): RegExp =>
  new RegExp(`\\.(${documentApps.find((app) => app.id === id)!.extensions.join('|')})$`, 'i')

// shared by the open route, the open-dialog filter and the rename gate so they cannot drift
// .odt/.odp (and the .ott/.ots/.otp templates) open through the docs/slides/sheets
// import filters (converted copies)
const DOCX_RE = extensionRe('docs')
const XLSX_RE = extensionRe('sheets')
const PPTX_RE = extensionRe('slides')
const PDF_RE = extensionRe('pdf')
const TEXT_RE = extensionRe('markdown')
const HTML_RE = extensionRe('html')

export { DOCX_RE, XLSX_RE, PPTX_RE, PDF_RE, HTML_RE, TEXT_RE }

/** the app that owns files with this path's extension; the registry is the source of truth */
export { appForPath }

/** A legal name in an extension no app routes to would turn an openable file into an unopenable one. */
export function renameStaysInApp(from: string, newName: string): boolean {
  const current = appForPath(from)
  if (!current) return true
  return appForPath(newName) === current
}

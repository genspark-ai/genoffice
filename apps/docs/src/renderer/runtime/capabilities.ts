import type { DocsErrorCode } from '../../shared/operation-result'
import type { RuntimeDocument } from './docs-runtime'
import { DOCS_READ_TOOLS } from '../../shared/tool-catalog'

const commentTools = new Set(['add_comment', 'reply_comment', 'resolve_comment', 'delete_comment'])
const noteTools = new Set([
  'read_notes',
  'insert_footnote',
  'insert_endnote',
  'delete_note',
  'edit_note',
])

/** The same availability decision is used by discovery and execution. */
export function toolUnavailable(
  name: string,
  doc: RuntimeDocument,
): { reason: string; errorCode: DocsErrorCode } | undefined {
  if (
    !doc.editor.isEditable &&
    !DOCS_READ_TOOLS.has(name) &&
    name !== 'stop_document_draft' &&
    !(commentTools.has(name) && doc.commentsAllowed)
  )
    return { reason: 'The document is read-only', errorCode: 'permission_denied' }
  if ((commentTools.has(name) || name === 'read_comments') && !doc.comments)
    return {
      reason: 'Comments are not available in this document session',
      errorCode: 'unavailable',
    }
  if (noteTools.has(name) && !doc.notes)
    return { reason: 'Notes are not available in this document session', errorCode: 'unavailable' }
  if (name === 'set_header_footer' && !doc.headerFooter)
    return {
      reason: 'Header and footer editing is not available in this document session',
      errorCode: 'unavailable',
    }
  if ((name === 'set_page_setup' || name === 'insert_section_break') && !doc.pageSetup)
    return {
      reason: 'Page setup is not available in this document session',
      errorCode: 'unavailable',
    }
  if ((name === 'list_styles' || name === 'define_style') && !doc.extras?.styles)
    return { reason: 'Styles are not available in this document session', errorCode: 'unavailable' }
  if (name === 'set_watermark' && !doc.extras?.watermark)
    return {
      reason: 'Watermarks are not available in this document session',
      errorCode: 'unavailable',
    }
  if (name === 'read_document_diagnostics' && !doc.annotations)
    return {
      reason: 'Host diagnostics are not available in this document session',
      errorCode: 'unavailable',
    }
  if ((name === 'accept_changes' || name === 'reject_changes') && doc.extras?.revisionsLocked)
    return {
      reason: 'Document protection forbids revision decisions',
      errorCode: 'permission_denied',
    }
}

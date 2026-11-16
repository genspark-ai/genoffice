import type { Editor } from '@tiptap/core'
import type { DocsEngineTarget } from '../../shared/document-query'
import type { DocsAnnotations } from '../sdk/annotations'

export interface DiagnosticsQuery {
  source?: string
  limit?: number
  cursor?: string
}

interface DiagnosticsCursor {
  documentId: string
  revision: number
  snapshotDocumentId: string
  snapshotRevision: number
  version: number
  source?: string
  limit: number
  offset: number
}

type DiagnosticItem = {
  kind: 'diagnostic'
  source: string
  code?: string
  severity: 'info' | 'warning' | 'error'
  message: string
  messageComplete: boolean
  from: number
  to: number
  stale: boolean
  checkedReference: { documentId: string; revision: number }
  targets?: DocsEngineTarget[]
  targetsComplete?: boolean
}

type SourceItem = {
  kind: 'source'
  source: string
  status: 'pending' | 'complete' | 'stale' | 'failed'
  checkedReference: { documentId: string; revision: number }
  coverage: { kind: 'document-text' | 'ranges'; from: number; to: number; rangeCount?: number }
  error?: string
}

/** Snapshot-only read: starting or refreshing a host checker remains a host action. */
export function queryDocumentDiagnostics(
  editor: Editor,
  annotations: DocsAnnotations,
  documentId: string,
  revision: number,
  input: DiagnosticsQuery,
) {
  const snapshot = annotations.snapshot()
  let source = input.source
  let limit = input.limit ?? 20
  let offset = 0
  if (input.cursor) {
    if (Object.keys(input).some((key) => key !== 'cursor'))
      throw new Error('A diagnostic continuation cannot change query filters or budgets')
    let cursor: DiagnosticsCursor
    try {
      cursor = JSON.parse(decodeURIComponent(atob(input.cursor))) as DiagnosticsCursor
    } catch {
      throw new Error('Invalid diagnostic continuation')
    }
    if (
      cursor.documentId !== documentId ||
      cursor.revision !== revision ||
      cursor.snapshotDocumentId !== snapshot.documentId ||
      cursor.snapshotRevision !== snapshot.revision ||
      cursor.version !== snapshot.version
    )
      throw new Error('Stale diagnostic continuation; read_document_diagnostics again.')
    source = cursor.source
    limit = cursor.limit
    offset = cursor.offset
  }
  if (
    !Number.isSafeInteger(limit) ||
    limit < 1 ||
    limit > 50 ||
    !Number.isSafeInteger(offset) ||
    offset < 0 ||
    (source !== undefined && (typeof source !== 'string' || source.length > 1000))
  )
    throw new Error('Invalid diagnostic query window')

  const items: Array<SourceItem | DiagnosticItem> = []
  const sourceStates = new Map(snapshot.sources.map((entry) => [entry.source, entry]))
  let withheldDiagnostics = 0
  const isWithheld = (from: number, to: number) => {
    let withheld = false
    editor.state.doc.nodesBetween(from, to, (node) => {
      if (node.attrs.redact || node.marks.some((mark) => mark.type.name === 'redaction'))
        withheld = true
    })
    return withheld
  }
  for (const state of snapshot.sources) {
    if (source !== undefined && source !== state.source) continue
    items.push({
      kind: 'source',
      source: state.source.slice(0, 1000),
      status: state.status,
      checkedReference: { documentId: state.documentId, revision: state.revision },
      coverage: state.coverage,
      ...(state.error ? { error: state.error.slice(0, 2000) } : {}),
    })
  }
  for (const finding of snapshot.diagnostics) {
    if (source !== undefined && source !== finding.source) continue
    if (isWithheld(finding.from, finding.to)) {
      withheldDiagnostics++
      continue
    }
    const state = sourceStates.get(finding.source)
    const stale =
      !!finding.stale ||
      state?.status !== 'complete' ||
      finding.documentId !== snapshot.documentId ||
      finding.revision !== snapshot.revision
    const targets: DocsEngineTarget[] = []
    let targetsComplete = true
    if (!stale && items.length >= offset && items.length < offset + limit)
      editor.state.doc.forEach((node, position) => {
        if (position >= finding.to || position + node.nodeSize <= finding.from) return
        if (targets.length === 32) {
          targetsComplete = false
          return
        }
        targets.push({
          kind: 'block',
          documentId,
          revision,
          expectedType: node.type.name,
          range: { from: position, to: position + node.nodeSize },
        })
      })
    items.push({
      kind: 'diagnostic',
      source: finding.source.slice(0, 1000),
      ...(finding.code ? { code: finding.code.slice(0, 1000) } : {}),
      severity: finding.severity ?? 'warning',
      message: finding.message.slice(0, 2000),
      messageComplete: finding.message.length <= 2000,
      from: finding.from,
      to: finding.to,
      stale,
      checkedReference: { documentId: finding.documentId, revision: finding.revision },
      ...(!stale ? { targets, targetsComplete } : {}),
    })
  }
  const page: typeof items = []
  let characters = 0
  for (const item of items.slice(offset, offset + limit)) {
    const size = JSON.stringify(item).length
    if (characters + size > 18000 && page.length) break
    characters += size
    page.push(item)
  }
  const nextOffset = offset + page.length
  const complete = nextOffset >= items.length
  return {
    documentId,
    revision,
    diagnosticVersion: snapshot.version,
    checksComplete: snapshot.sources
      .filter((entry) => source === undefined || source === entry.source)
      .every((entry) => entry.status === 'complete'),
    withheldDiagnostics,
    page: {
      complete,
      ...(!complete
        ? {
            nextCursor: btoa(
              encodeURIComponent(
                JSON.stringify({
                  documentId,
                  revision,
                  snapshotDocumentId: snapshot.documentId,
                  snapshotRevision: snapshot.revision,
                  version: snapshot.version,
                  source,
                  limit,
                  offset: nextOffset,
                } satisfies DiagnosticsCursor),
              ),
            ),
          }
        : {}),
    },
    items: page,
  }
}

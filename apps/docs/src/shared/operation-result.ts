import type { ToolExecution } from '@genoffice/agent-core'

export const DOCS_ERROR_CODES = [
  'invalid_input',
  'stale_target',
  'permission_denied',
  'unavailable',
  'canceled',
  'host_failure',
  'execution_failure',
] as const
export type DocsErrorCode = (typeof DOCS_ERROR_CODES)[number]

export interface DocsImageReferenceIssue {
  reference: string
  reason: 'unbound' | 'malformed' | 'expired' | 'unsupported' | 'missing'
}

export interface DocsOperationResult extends ToolExecution {
  receipt?: {
    outcome:
      | 'rejected'
      | 'preview'
      | 'read'
      | 'no-op'
      | 'applied'
      | 'failure-after-commit'
      | 'independent-effect'
    revisionBefore: number
    revisionAfter: number
    persistence: 'unchanged' | 'pending' | 'completed' | 'unknown'
    resolvedTargets?: import('./document-query').DocsEngineTarget[]
    operations?: Array<{
      op: string
      matched: number
      changed: number
      matchCount?: number
      changeCount?: number
      structuralEffects?: Array<{
        reference: string
        kind: string
        effect: 'removed' | 'marked-for-deletion'
      }>
    }>
  }
  documentId: string
  revision: number
  /** null means the host lost the IO response and the file effect is unknown. */
  committed: boolean | null
  pending?: boolean
  errorCode?: DocsErrorCode
  /** Exact matching/changing occurrences or cells, when provided by the operation. */
  matchCount?: number
  changeCount?: number
  media?: { data: string; mimeType: string }
  imageReferenceErrors?: DocsImageReferenceIssue[]
}

export interface DocsSaveOutcome {
  ok: boolean
  path?: string
  error?: string
  errorCode?: DocsErrorCode
  committed: boolean | null
  pending: boolean
}

export function isDocsSaveOutcome(value: unknown): value is DocsSaveOutcome {
  if (!value || typeof value !== 'object') return false
  const result = value as Record<string, unknown>
  return (
    typeof result.ok === 'boolean' &&
    typeof result.pending === 'boolean' &&
    (typeof result.committed === 'boolean' || result.committed === null) &&
    (result.path === undefined || typeof result.path === 'string') &&
    (result.error === undefined || typeof result.error === 'string') &&
    (result.errorCode === undefined || DOCS_ERROR_CODES.some((code) => code === result.errorCode))
  )
}

export type DocsSaveResult = DocsOperationResult & DocsSaveOutcome

/** Carries the standard tool-execution error flag and its application diagnostics across host transports. */
export class DocsOperationError extends Error {
  constructor(
    message: string,
    readonly result: DocsOperationResult,
  ) {
    super(message)
  }
}

export function isDocsOperationResult(value: unknown): value is DocsOperationResult {
  if (!value || typeof value !== 'object') return false
  const result = value as Record<string, unknown>
  return (
    typeof result.documentId === 'string' &&
    typeof result.revision === 'number' &&
    Number.isSafeInteger(result.revision) &&
    result.revision >= 0 &&
    (typeof result.committed === 'boolean' || result.committed === null) &&
    (result.pending === undefined || typeof result.pending === 'boolean') &&
    typeof result.output === 'string' &&
    typeof result.summary === 'string' &&
    (result.isError === undefined || typeof result.isError === 'boolean') &&
    (result.mutated === undefined || typeof result.mutated === 'boolean') &&
    (result.media === undefined ||
      (typeof result.media === 'object' &&
        result.media !== null &&
        'data' in result.media &&
        typeof result.media.data === 'string' &&
        'mimeType' in result.media &&
        typeof result.media.mimeType === 'string')) &&
    (result.errorCode === undefined || DOCS_ERROR_CODES.some((code) => code === result.errorCode))
  )
}

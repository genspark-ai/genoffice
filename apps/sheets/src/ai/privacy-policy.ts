import type { CellScalar } from '../domain/workbook.types'

export type UploadPolicy = 'allow' | 'redact' | 'statistics-only' | 'deny'

export interface ContextColumn {
  readonly name: string
  readonly policy: UploadPolicy
  readonly values: readonly CellScalar[]
}

export interface SafeContextColumn {
  readonly name: string
  readonly policy: UploadPolicy
  readonly values?: readonly CellScalar[]
  readonly statistics?: {
    readonly count: number
    readonly nonEmptyCount: number
  }
}

export function applyPrivacyPolicy(
  columns: readonly ContextColumn[],
): readonly SafeContextColumn[] {
  const safeColumns: SafeContextColumn[] = []
  for (const column of columns) {
    if (column.policy === 'deny') continue
    if (column.policy === 'statistics-only') {
      safeColumns.push({
        name: column.name,
        policy: column.policy,
        statistics: {
          count: column.values.length,
          nonEmptyCount: column.values.filter((value) => value !== null && value !== '').length,
        },
      })
      continue
    }
    safeColumns.push({
      name: column.name,
      policy: column.policy,
      values: column.policy === 'redact' ? column.values.map(redactScalar) : column.values,
    })
  }
  return safeColumns
}

function redactScalar(value: CellScalar): CellScalar {
  if (typeof value !== 'string') return value
  return value
    .replace(/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi, '[EMAIL]')
    .replace(/(?<!\d)(?:\+?\d[\d ()-]{7,}\d)(?!\d)/g, '[PHONE]')
}

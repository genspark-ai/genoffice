import type { WorkbookCommandBatch } from '../domain/workbook-dsl'

export class UnsupportedPromptError extends Error {
  constructor() {
    super('Try “set A1 to 42”, “formula B1 = SUM(A1:A10)”, or “rename sheet to Budget”.')
    this.name = 'UnsupportedPromptError'
  }
}

export function planPrompt(
  prompt: string,
  options: { readonly revision: number; readonly sheetId: string },
): WorkbookCommandBatch {
  const transactionId = `local-${crypto.randomUUID()}`
  const normalized = prompt.trim()
  const setMatch = /^set\s+([a-z]{1,3}[1-9][0-9]{0,6})\s+to\s+(.+)$/i.exec(normalized)
  if (setMatch) {
    const address = setMatch[1]
    const rawValue = setMatch[2]
    if (!address || !rawValue) {
      throw new UnsupportedPromptError()
    }
    return {
      dslVersion: 1,
      transactionId,
      baseRevision: options.revision,
      summary: `Set ${address.toUpperCase()}`,
      operations: [
        {
          op: 'set_cell',
          sheetId: options.sheetId,
          address: address.toUpperCase(),
          value: parseScalar(rawValue),
        },
      ],
    }
  }

  const formulaMatch = /^formula\s+([a-z]{1,3}[1-9][0-9]{0,6})\s*=\s*(.+)$/i.exec(normalized)
  if (formulaMatch) {
    const address = formulaMatch[1]
    const formula = formulaMatch[2]
    if (!address || !formula) {
      throw new UnsupportedPromptError()
    }
    return {
      dslVersion: 1,
      transactionId,
      baseRevision: options.revision,
      summary: `Set formula in ${address.toUpperCase()}`,
      operations: [
        {
          op: 'set_formula',
          sheetId: options.sheetId,
          address: address.toUpperCase(),
          formula: `=${formula}`,
        },
      ],
    }
  }

  const renameMatch = /^rename\s+sheet\s+to\s+(.+)$/i.exec(normalized)
  if (renameMatch?.[1]) {
    return {
      dslVersion: 1,
      transactionId,
      baseRevision: options.revision,
      summary: 'Rename current sheet',
      operations: [
        {
          op: 'rename_sheet',
          sheetId: options.sheetId,
          name: renameMatch[1].trim(),
        },
      ],
    }
  }

  throw new UnsupportedPromptError()
}

function parseScalar(rawValue: string): string | number | boolean | null {
  if (rawValue === 'null') return null
  if (rawValue.toLowerCase() === 'true') return true
  if (rawValue.toLowerCase() === 'false') return false
  const numericValue = Number(rawValue)
  return Number.isFinite(numericValue) && rawValue.trim() !== '' ? numericValue : rawValue
}

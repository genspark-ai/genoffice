export interface CellCoordinates {
  readonly row: number
  readonly column: number
}

export interface RangeBounds {
  readonly startRow: number
  readonly startColumn: number
  readonly endRow: number
  readonly endColumn: number
}

export function parseAddress(address: string): CellCoordinates {
  // $-anchored A1 notation is equivalent here; some producers store pivot
  // location refs as $C$33 and refreshing such a pivot must not choke.
  const match = /^\$?([A-Z]+)\$?([1-9][0-9]*)$/.exec(address)
  if (!match?.[1] || !match[2]) throw new Error(`Invalid cell address: ${address}`)
  let column = 0
  for (const character of match[1]) {
    column = column * 26 + character.charCodeAt(0) - 64
  }
  return { row: Number(match[2]) - 1, column: column - 1 }
}

/// Inverse of parseAddress's column parsing: 0 → A, 25 → Z, 26 → AA.
export function columnLabel(column: number): string {
  let label = ''
  let remaining = column + 1
  while (remaining > 0) {
    remaining -= 1
    label = String.fromCharCode(65 + (remaining % 26)) + label
    remaining = Math.floor(remaining / 26)
  }
  return label
}

export function formatAddress(row: number, column: number): string {
  return `${columnLabel(column)}${row + 1}`
}

/// A → 0, Z → 25, AA → 26.
export function columnIndex(label: string): number {
  if (!/^[A-Z]+$/.test(label)) throw new Error(`Invalid column label: ${label}`)
  let column = 0
  for (const character of label) {
    column = column * 26 + character.charCodeAt(0) - 64
  }
  return column - 1
}

/// Accepts "A1:C10" or a single cell "B2"; normalizes so start ≤ end.
export function parseRange(range: string): RangeBounds {
  const parts = range.split(':')
  if (parts.length > 2 || !parts[0]) throw new Error(`Invalid range: ${range}`)
  const first = parseAddress(parts[0])
  const second = parts[1] ? parseAddress(parts[1]) : first
  return {
    startRow: Math.min(first.row, second.row),
    startColumn: Math.min(first.column, second.column),
    endRow: Math.max(first.row, second.row),
    endColumn: Math.max(first.column, second.column),
  }
}

export function rangeCellCount(bounds: RangeBounds): number {
  return (bounds.endRow - bounds.startRow + 1) * (bounds.endColumn - bounds.startColumn + 1)
}

export function rangeAddresses(bounds: RangeBounds): string[] {
  const addresses: string[] = []
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      addresses.push(formatAddress(row, column))
    }
  }
  return addresses
}

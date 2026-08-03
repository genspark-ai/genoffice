import type { SelectionFormat } from './selection-format'

export type TriState = '' | 'on' | 'off'

/// Everything the Format Cells dialog edits, prefilled from the selection so
/// it opens showing the current format.
export interface FormatCellsDraft {
  readonly pattern: string
  readonly family: string
  readonly size: string
  readonly bold: TriState
  readonly italic: TriState
  readonly underline: TriState
  readonly strike: TriState
  readonly fontColor: string
  readonly border: string
  readonly borderColor: string
  /// OOXML ST_BorderStyle name applied with the preset.
  readonly borderStyle: string
  readonly fill: string
  readonly noFill: boolean
  readonly hAlign: string
  readonly vAlign: string
  readonly wrapText: TriState
  /// Ribbon rotate preset keys ('' = unchanged, '0' = clear).
  readonly rotation: string
  /// '' = unchanged, otherwise the indent step count as text.
  readonly indent: string
  readonly locked: TriState
  readonly hidden: TriState
}

export function draftFromSelection(format: SelectionFormat | null): FormatCellsDraft {
  const toTriState = (flag: boolean | undefined): TriState =>
    flag === undefined ? '' : flag ? 'on' : 'off'
  return {
    pattern: format?.numberFormat ?? '',
    family: format?.fontFamily ?? '',
    size: format?.fontSize != null ? String(format.fontSize) : '',
    bold: toTriState(format?.bold),
    italic: toTriState(format?.italic),
    underline: toTriState(format?.underline),
    strike: toTriState(format?.strike),
    // Native color inputs report lowercase hex; store the same casing so
    // "user picked the value it already had" compares equal.
    fontColor: format?.fontColor?.toLowerCase() ?? '',
    border: '',
    borderColor: '#000000',
    borderStyle: 'thin',
    fill: format?.fillColor?.toLowerCase() ?? '',
    noFill: false,
    hAlign: format?.horizontalAlignment ?? '',
    vAlign: format?.verticalAlignment === 'center' ? 'middle' : (format?.verticalAlignment ?? ''),
    wrapText: toTriState(format?.wrap),
    rotation: rotationPreset(format?.textRotation),
    // 0 echoes as '' so an untouched dialog never emits an indent command.
    indent: format?.indent ? String(format.indent) : '',
    locked: '',
    hidden: '',
  }
}

/// OOXML rotation value → the ribbon preset key it corresponds to; values
/// between presets echo as unchanged.
function rotationPreset(value: number | null | undefined): string {
  if (value == null) return ''
  const presets: Record<number, string> = {
    45: '45',
    135: '-45',
    90: '90',
    180: '-90',
    255: 'vertical',
  }
  return presets[value] ?? ''
}

/// Ribbon command strings for exactly the settings the user changed relative
/// to the prefilled draft — untouched settings never emit.
export function formatCellsCommands(initial: FormatCellsDraft, edited: FormatCellsDraft): string[] {
  const commands: string[] = []
  if (edited.pattern && edited.pattern !== initial.pattern) {
    commands.push(`format:${edited.pattern}`)
  }
  if (edited.family && edited.family !== initial.family) {
    commands.push(`font-family:${edited.family}`)
  }
  if (edited.size && edited.size !== initial.size) commands.push(`font-size:${edited.size}`)
  for (const key of ['bold', 'italic', 'underline', 'strike'] as const) {
    if (edited[key] && edited[key] !== initial[key]) commands.push(`${key}:${edited[key]}`)
  }
  if (edited.fontColor && edited.fontColor !== initial.fontColor) {
    commands.push(`font-color:${edited.fontColor}`)
  }
  // Borders have no echo (presets cannot represent arbitrary per-edge state),
  // so any picked preset is a change.
  if (edited.border) {
    commands.push(`border:${edited.border}:${edited.borderColor}:${edited.borderStyle}`)
  }
  if (edited.noFill) {
    if (initial.fill) commands.push('fill:none')
  } else if (edited.fill && edited.fill !== initial.fill) {
    commands.push(`fill:${edited.fill}`)
  }
  if (edited.hAlign && edited.hAlign !== initial.hAlign) {
    commands.push(`align:${edited.hAlign}`)
  }
  if (edited.vAlign && edited.vAlign !== initial.vAlign) {
    commands.push(`valign:${edited.vAlign}`)
  }
  if (edited.wrapText && edited.wrapText !== initial.wrapText) {
    commands.push(`wrap:${edited.wrapText}`)
  }
  if (edited.rotation && edited.rotation !== initial.rotation) {
    commands.push(`rotate:${edited.rotation}`)
  }
  if (edited.indent && edited.indent !== initial.indent) {
    commands.push(`indent:${edited.indent}`)
  }
  if (
    (edited.locked && edited.locked !== initial.locked) ||
    (edited.hidden && edited.hidden !== initial.hidden)
  ) {
    const locked =
      edited.locked && edited.locked !== initial.locked ? `locked-${edited.locked}` : ''
    const hidden =
      edited.hidden && edited.hidden !== initial.hidden ? `hidden-${edited.hidden}` : ''
    commands.push(`cellprot:${locked}:${hidden}`)
  }
  return commands
}

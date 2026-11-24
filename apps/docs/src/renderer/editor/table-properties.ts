import type { Node as PmNode } from '@tiptap/pm/model'
import type { Command, EditorState, Transaction } from '@tiptap/pm/state'
import { isInTable, selectedRect } from '@tiptap/pm/tables'
import type { TableAutoFitMode, TableLook } from '@genoffice/docx-engine'

interface SelectedTable {
  node: PmNode
  pos: number
  tableStart: number
  rect: ReturnType<typeof selectedRect>
}

function selectedTable(state: EditorState): SelectedTable | null {
  if (!isInTable(state)) return null
  try {
    const rect = selectedRect(state)
    const pos = rect.tableStart - 1
    const node = state.doc.nodeAt(pos)
    return node?.type.name === 'docTable' ? { node, pos, tableStart: rect.tableStart, rect } : null
  } catch {
    return null
  }
}

export function updateSelectedTableAttrs(patch: Record<string, unknown>): Command {
  return (state, dispatch) => {
    const selected = selectedTable(state)
    if (!selected) return false
    dispatch?.(
      state.tr.setNodeMarkup(selected.pos, undefined, {
        ...selected.node.attrs,
        ...patch,
      }),
    )
    return true
  }
}

function fixedGridWidths(table: PmNode, columnCount: number, contentWidthPx: number): number[] {
  const firstRow = table.firstChild
  const explicit: number[] = []
  firstRow?.forEach((cell) => {
    const widths = cell.attrs.colwidth as number[] | null
    if (widths?.length) explicit.push(...widths)
  })
  if (explicit.length === columnCount && explicit.every((width) => width > 0)) return explicit
  const pct = table.attrs.colWidthsPct as number[] | null
  if (pct?.length === columnCount) {
    const total = pct.reduce((sum, width) => sum + Math.max(0, width), 0) || 100
    return pct.map((width) =>
      Math.max(24, Math.round((Math.max(0, width) / total) * contentWidthPx)),
    )
  }
  return Array.from({ length: columnCount }, () =>
    Math.max(24, Math.floor(contentWidthPx / Math.max(1, columnCount))),
  )
}

/**
 * Word's AutoFit modes:
 * - contents: auto layout with no preferred table width
 * - window: auto layout with a 100% preferred width
 * - fixed: fixed layout using the current grid
 */
export function setTableAutoFit(mode: TableAutoFitMode, contentWidthPx: number): Command {
  return (state, dispatch) => {
    const selected = selectedTable(state)
    if (!selected) return false
    let tr: Transaction = state.tr
    const { node: table, rect, tableStart } = selected
    const tablePatch: Record<string, unknown> = {
      tblAutoFit: mode,
      tblAutoFitEdited: true,
    }

    if (mode === 'contents') {
      tablePatch.widthPx = null
      tablePatch.widthPct = null
      tablePatch.colWidthsPct = null
      table.descendants((node, pos) => {
        if (node.type.name !== 'docTableCell' && node.type.name !== 'docTableHeader') return
        if (node.attrs.colwidth) {
          tr = tr.setNodeMarkup(tableStart + pos, undefined, { ...node.attrs, colwidth: null })
        }
      })
    } else if (mode === 'window') {
      tablePatch.widthPx = null
      tablePatch.widthPct = 100
      // Word's AutoFit Window also drops w:tblInd: the table spans the text column
      // from margin to margin, whatever indent it had (probe 2026-11-23, tblInd -108)
      tablePatch.indentTwips = null
      tablePatch.tblIndentEdited = true
    } else {
      const widths = fixedGridWidths(table, rect.map.width, contentWidthPx)
      const total = widths.reduce((sum, width) => sum + width, 0)
      tablePatch.widthPx = total
      tablePatch.widthPct = null
      tablePatch.colWidthsPct = widths.map((width) => (width / total) * 100)
      table.descendants((node, pos) => {
        if (node.type.name !== 'docTableCell' && node.type.name !== 'docTableHeader') return
        try {
          const cell = rect.map.findCell(pos)
          tr = tr.setNodeMarkup(tableStart + pos, undefined, {
            ...node.attrs,
            colwidth: widths.slice(cell.left, cell.right),
          })
        } catch {
          // A malformed grid should not prevent the table-level mode switch.
        }
      })
    }

    tr = tr.setNodeMarkup(selected.pos, undefined, { ...table.attrs, ...tablePatch })
    dispatch?.(tr)
    return true
  }
}

export interface RepeatHeaderState {
  enabled: boolean
  active: boolean
}

export function repeatHeaderState(state: EditorState): RepeatHeaderState {
  const selected = selectedTable(state)
  if (!selected || selected.rect.top !== 0) return { enabled: false, active: false }
  let active = true
  for (let row = 0; row < selected.rect.bottom; row++) {
    if (!selected.node.child(row).attrs.repeatHeader) active = false
  }
  return { enabled: true, active }
}

/** Toggle the selected leading row(s) as repeating table headers. */
export function toggleRepeatHeaderRows(): Command {
  return (state, dispatch) => {
    const selected = selectedTable(state)
    if (!selected || selected.rect.top !== 0) return false
    const active = repeatHeaderState(state).active
    let tr = state.tr
    selected.node.forEach((row, offset, index) => {
      if (index >= selected.rect.bottom) return
      tr = tr.setNodeMarkup(selected.tableStart + offset, undefined, {
        ...row.attrs,
        repeatHeader: !active,
        repeatHeaderEdited: true,
      })
    })
    dispatch?.(tr)
    return true
  }
}

export function setTableLookOption(key: keyof TableLook, value: boolean): Command {
  return (state, dispatch) => {
    const selected = selectedTable(state)
    if (!selected) return false
    const current = (selected.node.attrs.tblLook as TableLook | null) ?? DEFAULT_TABLE_LOOK
    dispatch?.(
      state.tr.setNodeMarkup(selected.pos, undefined, {
        ...selected.node.attrs,
        tblLook: { ...current, [key]: value },
        tblLookEdited: true,
      }),
    )
    return true
  }
}

export interface TablePreset {
  styleId: string
  /** the style's name in the file (w:name) */
  name: string
  headerFill: string | null
  band1Fill: string | null
  band2Fill: string | null
  borderColor: string
}

const DEFAULT_TABLE_LOOK: TableLook = {
  firstRow: true,
  lastRow: false,
  firstColumn: true,
  lastColumn: false,
  bandedRows: true,
  bandedColumns: false,
}

/** the w:style of type table a preset is saved as */
export function tablePresetStyleXml(preset: TablePreset): string {
  const line = (side: string) =>
    `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="${preset.borderColor}"/>`
  const shaded = (type: string, fill: string | null) =>
    fill
      ? `<w:tblStylePr w:type="${type}"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${fill}"/></w:tcPr></w:tblStylePr>`
      : ''
  return (
    `<w:style w:type="table" w:customStyle="1" w:styleId="${preset.styleId}">` +
    `<w:name w:val="${preset.name}"/><w:uiPriority w:val="99"/>` +
    '<w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblInd w:w="0" w:type="dxa"/>' +
    `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(line).join('')}</w:tblBorders>` +
    '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>' +
    '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
    shaded('firstRow', preset.headerFill) +
    shaded('band1Horz', preset.band1Fill) +
    shaded('band2Horz', preset.band2Fill) +
    '</w:style>'
  )
}

/**
 * Gives the table the preset's style, which the caller defines in styles.xml. The shading comes
 * from the style; the border colour is also set on the cells, where the page draws it from.
 */
export function applyTablePreset(preset: TablePreset): Command {
  return (state, dispatch) => {
    const selected = selectedTable(state)
    if (!selected) return false
    const border = { style: 'single', szEighths: 4, color: preset.borderColor }
    const borders = { top: border, right: border, bottom: border, left: border }
    let tr = state.tr
    selected.node.forEach((row, rowOffset) => {
      row.forEach((cell, cellOffset) => {
        tr = tr.setNodeMarkup(selected.tableStart + rowOffset + 1 + cellOffset, undefined, {
          ...cell.attrs,
          fill: null,
          borders,
        })
      })
    })
    tr = tr.setNodeMarkup(selected.pos, undefined, {
      ...selected.node.attrs,
      tblStyleId: preset.styleId,
      // without a w:tblLook Word applies none of the style's conditional formatting
      tblLook: (selected.node.attrs.tblLook as TableLook | null) ?? DEFAULT_TABLE_LOOK,
      tblLookEdited: true,
    })
    dispatch?.(tr)
    return true
  }
}

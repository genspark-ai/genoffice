import { useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { Command } from '@tiptap/pm/state'
import { isInTable, selectedRect } from '@tiptap/pm/tables'
import type { TableAutoFitMode } from '@genoffice/docx-engine'
import { useI18n, type StringKey } from '../i18n/locale'
import { setTableAutoFit, updateSelectedTableAttrs } from '../editor/table-properties'
import { setSelectedColumnPreferredWidth } from '../editor/table-sizing'
import {
  editableOrToast,
  moveCaretInTable,
  setCellProps,
  setRowProps,
  type CellVAlign,
} from '../editor/table-ops'
import { useModalKeys } from './modal-keys'
import { LengthInput } from './LengthInput'
import { UNIT_NAME_KEYS, useMeasurementUnit } from '../use-measurement'
import { IconTableProperties } from './icons'
import { Unavailable } from './Unavailable'

export const TABLE_AUTO_FIT_OPTIONS: Array<[TableAutoFitMode, StringKey]> = [
  ['contents', 'ribbonAutoFitContents'],
  ['window', 'ribbonAutoFitWindow'],
  ['fixed', 'ribbonFixedColumnWidth'],
]

export type TablePropertiesTab = 'table' | 'row' | 'column' | 'cell' | 'alt'

const TWIPS_PER_CM = 567
const PX_PER_CM = 96 / 2.54
const TABS: TablePropertiesTab[] = ['table', 'row', 'column', 'cell', 'alt']
const TAB_LABELS: Record<TablePropertiesTab, StringKey> = {
  table: 'ribbonTableData',
  row: 'ribbonTabRow',
  column: 'ribbonTabColumn',
  cell: 'ribbonTabCell',
  alt: 'ribbonTabAltText',
}

type TableWrapMode = 'none' | 'left' | 'right'
type TableAlign = 'left' | 'center' | 'right'

export interface TableValue {
  autoFit: TableAutoFitMode
  align: TableAlign
  wrap: TableWrapMode
  floatSuppressed: boolean
  positionXCm: number
  positionYCm: number
  distanceCm: number
  marginTopCm: number
  marginRightCm: number
  marginBottomCm: number
  marginLeftCm: number
}

const cmOf = (value: unknown, fallback = 0) =>
  +((Number.isFinite(Number(value)) ? Number(value) : fallback) / TWIPS_PER_CM).toFixed(2)

export function tableValueOf(attrs: Record<string, unknown>): TableValue {
  const margins = (attrs.cellMar as Record<string, number> | null) ?? {}
  const distance = (attrs.tblFloatDistance as Record<string, number> | null) ?? {}
  const wrap: TableWrapMode =
    attrs.tblFloat === 'left' || attrs.tblFloat === 'right'
      ? attrs.tblFloat
      : attrs.tblFloatSource === 'left' || attrs.tblFloatSource === 'right'
        ? attrs.tblFloatSource
        : 'none'
  return {
    autoFit:
      attrs.tblAutoFit === 'contents' || attrs.tblAutoFit === 'window' ? attrs.tblAutoFit : 'fixed',
    align: attrs.tblAlign === 'center' || attrs.tblAlign === 'right' ? attrs.tblAlign : 'left',
    wrap,
    floatSuppressed: attrs.tblFloatSuppressed === true,
    positionXCm: cmOf(attrs.tblFloatXTwips),
    positionYCm: cmOf(attrs.tblFloatYTwips),
    distanceCm: cmOf(distance.right ?? distance.left ?? distance.top ?? distance.bottom, 180),
    marginTopCm: cmOf(margins.top),
    marginRightCm: cmOf(margins.right, 108),
    marginBottomCm: cmOf(margins.bottom),
    marginLeftCm: cmOf(margins.left, 108),
  }
}

type TableGroup = 'autoFit' | 'align' | 'wrap' | 'margins'

/** writes only the touched groups; the width measurement mirrors the ribbon's AutoFit path */
export function applyTableValue(
  editor: Editor,
  value: TableValue,
  touched: ReadonlySet<TableGroup>,
  sectionContentWidthPx: number,
): void {
  if (!isInTable(editor.state) || touched.size === 0) return
  const run = (command: Command) => {
    editor.view.focus()
    command(editor.state, editor.view.dispatch)
  }
  const tableAttrs = editor.getAttributes('docTable')
  if (touched.has('autoFit')) run(setTableAutoFit(value.autoFit, sectionContentWidthPx))
  const patch: Record<string, unknown> = {}
  if (touched.has('align')) patch.tblAlign = value.align
  if (touched.has('margins')) {
    const twips = (cm: number) => Math.max(0, Math.round(cm * TWIPS_PER_CM))
    patch.cellMar = {
      top: twips(value.marginTopCm),
      right: twips(value.marginRightCm),
      bottom: twips(value.marginBottomCm),
      left: twips(value.marginLeftCm),
    }
    patch.cellMarEdited = true
  }
  if (touched.has('wrap')) {
    let measuredWidthPx = Number(tableAttrs.widthPx) || 0
    if (!(measuredWidthPx > 0)) {
      try {
        const rect = selectedRect(editor.state)
        const tableDom = editor.view.nodeDOM(rect.tableStart - 1)
        if (tableDom instanceof HTMLElement) {
          const zoomEl = document.querySelector('.doc-zoom') as HTMLElement | null
          const zoom = zoomEl ? parseFloat(getComputedStyle(zoomEl).zoom || '1') || 1 : 1
          measuredWidthPx = tableDom.getBoundingClientRect().width / zoom
        }
      } catch {
        // a malformed table falls back to the section width below
      }
    }
    measuredWidthPx = Math.max(1, Math.round(measuredWidthPx || sectionContentWidthPx))
    const positionedWidthPx =
      value.autoFit === 'window' ? Math.round(sectionContentWidthPx) : measuredWidthPx
    const twips = (cm: number) => Math.max(0, Math.round(cm * TWIPS_PER_CM))
    const signedTwips = (cm: number) => Math.round(cm * TWIPS_PER_CM)
    const floatX =
      value.wrap === 'right' && Math.abs(value.positionXCm) < 0.001
        ? value.autoFit === 'window'
          ? 0
          : Math.max(0, Math.round((sectionContentWidthPx - positionedWidthPx) * 15))
        : signedTwips(value.positionXCm)
    const keepFloatSuppressed = value.floatSuppressed && value.wrap !== 'none'
    Object.assign(patch, {
      tblFloatWidthPx: value.wrap === 'right' ? positionedWidthPx : null,
      tblFloat: keepFloatSuppressed ? null : value.wrap,
      tblFloatSource: value.wrap,
      tblFloatSuppressed: keepFloatSuppressed,
      tblFloatXTwips: value.wrap === 'none' ? null : floatX,
      tblFloatYTwips: value.wrap === 'none' ? null : signedTwips(value.positionYCm),
      tblFloatHorzAnchor:
        value.wrap === 'none' ? null : (tableAttrs.tblFloatHorzAnchor ?? 'margin'),
      tblFloatVertAnchor:
        value.wrap === 'none' ? null : (tableAttrs.tblFloatVertAnchor ?? 'margin'),
      tblFloatDistance:
        value.wrap === 'none'
          ? null
          : {
              top: twips(value.distanceCm),
              right: twips(value.distanceCm),
              bottom: twips(value.distanceCm),
              left: twips(value.distanceCm),
            },
      tblFloatEdited: true,
    })
  }
  if (Object.keys(patch).length) run(updateSelectedTableAttrs(patch))
}

export interface RowValue {
  specifyHeight: boolean
  heightCm: number
  rule: 'atLeast' | 'exact'
  allowBreak: boolean
  repeatHeader: boolean
}

/** only the touched row fields are written; the header flag needs the first row as well */
export function rowPatchFor(
  row: RowValue,
  touched: ReadonlySet<keyof RowValue>,
  firstRow: boolean,
): Record<string, unknown> {
  const patch: Record<string, unknown> = {}
  if (touched.has('specifyHeight') || touched.has('heightCm') || touched.has('rule')) {
    patch.heightTwips =
      row.specifyHeight && row.heightCm > 0 ? Math.round(row.heightCm * TWIPS_PER_CM) : null
    patch.heightRule = row.specifyHeight ? row.rule : null
  }
  if (touched.has('allowBreak')) {
    patch.cantSplit = !row.allowBreak
    patch.cantSplitEdited = true
  }
  if (touched.has('repeatHeader') && firstRow) {
    patch.repeatHeader = row.repeatHeader
    patch.repeatHeaderEdited = true
  }
  return patch
}

export interface PreferredWidth {
  cm: number
  /** share of the table, used when the width is measured in percent */
  pct: number
  inPct: boolean
  /** the whole grid, what a percent is a share of */
  tableCm: number
}

/** Measure in: the width on screen is restated in the unit switched to */
export function measureWidthIn(width: PreferredWidth, inPct: boolean): PreferredWidth {
  if (inPct === width.inPct || !(width.tableCm > 0)) return { ...width, inPct }
  return inPct
    ? { ...width, inPct, pct: +Math.min(100, (width.cm / width.tableCm) * 100).toFixed(1) }
    : { ...width, inPct, cm: +((width.pct / 100) * width.tableCm).toFixed(2) }
}

interface CellValue {
  width: PreferredWidth
  vAlign: CellVAlign
  sameMargins: boolean
  marginTopCm: number
  marginRightCm: number
  marginBottomCm: number
  marginLeftCm: number
  wrap: boolean
  /** w:tcFitText */
  fitText: boolean
}

interface Selection {
  top: number
  bottom: number
  left: number
  right: number
  rows: number
  cols: number
  row: RowValue
  width: PreferredWidth
  cell: CellValue
}

function selectionOf(editor: Editor, sectionContentWidthPx: number): Selection | null {
  if (!isInTable(editor.state)) return null
  const rect = selectedRect(editor.state)
  const rowNode = rect.table.child(rect.top)
  const cellPos = rect.map.map[rect.top * rect.map.width + rect.left]
  const cellNode = editor.state.doc.nodeAt(rect.tableStart + cellPos)
  const cellRect = rect.map.findCell(cellPos)
  const colwidth = cellNode?.attrs.colwidth as number[] | null
  const pct = rect.table.attrs.colWidthsPct as number[] | null
  const widthPx =
    colwidth?.[rect.left - cellRect.left] ??
    (pct?.length === rect.map.width ? (pct[rect.left] / 100) * sectionContentWidthPx : 0)
  const sharePct =
    pct?.length === rect.map.width
      ? (pct[rect.left] / pct.reduce((sum, p) => sum + p, 0)) * 100
      : 100 / rect.map.width
  const width: PreferredWidth = {
    cm: +(widthPx / PX_PER_CM).toFixed(2),
    pct: +sharePct.toFixed(1),
    inPct: !!cellNode?.attrs.widthPct,
    tableCm: sharePct > 0 ? widthPx / PX_PER_CM / (sharePct / 100) : 0,
  }
  const heightTwips = Number(rowNode.attrs.heightTwips) || 0
  const margins = (cellNode?.attrs.cellMar as Record<string, number> | null) ?? null
  const tableMargins = (rect.table.attrs.cellMar as Record<string, number> | null) ?? {}
  const vAlign = cellNode?.attrs.vAlign
  return {
    top: rect.top,
    bottom: rect.bottom,
    left: rect.left,
    right: rect.right,
    rows: rect.map.height,
    cols: rect.map.width,
    row: {
      specifyHeight: heightTwips > 0,
      heightCm: heightTwips > 0 ? cmOf(heightTwips) : 0,
      rule: rowNode.attrs.heightRule === 'exact' ? 'exact' : 'atLeast',
      allowBreak: !rowNode.attrs.cantSplit,
      repeatHeader: !!rowNode.attrs.repeatHeader,
    },
    width,
    cell: {
      width,
      vAlign: vAlign === 'center' || vAlign === 'bottom' ? vAlign : 'top',
      sameMargins: margins === null,
      marginTopCm: cmOf(margins?.top ?? tableMargins.top),
      marginRightCm: cmOf(margins?.right ?? tableMargins.right, 108),
      marginBottomCm: cmOf(margins?.bottom ?? tableMargins.bottom),
      marginLeftCm: cmOf(margins?.left ?? tableMargins.left, 108),
      wrap: !cellNode?.attrs.noWrap,
      fitText: !!cellNode?.attrs.fitText,
    },
  }
}

export function TablePropertiesDialog({
  editor,
  sectionContentWidthPx,
  initialTab = 'table',
  onClose,
}: {
  editor: Editor
  sectionContentWidthPx: number
  initialTab?: TablePropertiesTab
  onClose: () => void
}) {
  const { t } = useI18n()
  const [tab, setTab] = useState<TablePropertiesTab>(initialTab)
  const [tableValue, setTableValue] = useState(() => tableValueOf(editor.getAttributes('docTable')))
  const [tableTouched, setTableTouched] = useState<Set<TableGroup>>(() => new Set())
  // the row / column / cell tabs follow the caret (prev/next commit and move, like Word)
  const [sel, setSel] = useState(() => selectionOf(editor, sectionContentWidthPx))
  const [row, setRow] = useState<RowValue | null>(sel?.row ?? null)
  const [rowTouched, setRowTouched] = useState<Set<keyof RowValue>>(() => new Set())
  const [columnWidth, setColumnWidth] = useState<PreferredWidth | null>(sel?.width ?? null)
  const [columnTouched, setColumnTouched] = useState(false)
  const [cell, setCell] = useState<CellValue | null>(sel?.cell ?? null)
  const [cellTouched, setCellTouched] = useState<Set<keyof CellValue>>(() => new Set())
  // Word keeps cell margins, wrap and fit behind the Cell tab's Options button
  const [cellOptions, setCellOptions] = useState<{
    cell: CellValue
    touched: Set<keyof CellValue>
  } | null>(null)
  const cancelCellOptions = () => {
    if (!cellOptions) return
    setCell(cellOptions.cell)
    setCellTouched(cellOptions.touched)
    setCellOptions(null)
  }
  // Escape with the focus on neither dialog still only leaves Options
  const keys = useModalKeys(cellOptions ? cancelCellOptions : onClose)
  const unit = useMeasurementUnit()
  const tableAttrs = editor.getAttributes('docTable')
  const [altTitle, setAltTitle] = useState(String(tableAttrs.tblCaption ?? ''))
  const [altDescription, setAltDescription] = useState(String(tableAttrs.tblDescription ?? ''))
  const [altTouched, setAltTouched] = useState(false)

  const updateTable = (group: TableGroup, patch: Partial<TableValue>) => {
    setTableValue((current) => ({ ...current, ...patch }))
    setTableTouched((current) => new Set(current).add(group))
  }
  const updateRow = (patch: Partial<RowValue>) => {
    setRow((current) => (current ? { ...current, ...patch } : current))
    setRowTouched((current) => {
      const next = new Set(current)
      for (const key of Object.keys(patch) as Array<keyof RowValue>) next.add(key)
      return next
    })
  }
  const updateCell = (patch: Partial<CellValue>) => {
    setCell((current) => (current ? { ...current, ...patch } : current))
    setCellTouched((current) => {
      const next = new Set(current)
      for (const key of Object.keys(patch) as Array<keyof CellValue>) next.add(key)
      return next
    })
  }

  const run = (command: Command) => {
    editor.view.focus()
    command(editor.state, editor.view.dispatch)
  }
  const commitRow = () => {
    if (!row || rowTouched.size === 0) return
    const patch = rowPatchFor(row, rowTouched, sel?.top === 0)
    if (Object.keys(patch).length) run(setRowProps(patch))
    setRowTouched(new Set())
  }
  const commitWidth = (width: PreferredWidth) => {
    if (!((width.inPct ? width.pct : width.cm) > 0)) return
    run(
      setSelectedColumnPreferredWidth(
        width.inPct ? { pct: width.pct } : { px: Math.round(width.cm * PX_PER_CM) },
        sectionContentWidthPx,
      ),
    )
  }
  const commitColumn = () => {
    if (!columnTouched || !columnWidth) return
    commitWidth(columnWidth)
    setColumnTouched(false)
  }
  const commitCell = () => {
    if (!cell || cellTouched.size === 0) return
    if (cellTouched.has('width')) commitWidth(cell.width)
    const patch: Record<string, unknown> = {}
    if (cellTouched.has('vAlign')) patch.vAlign = cell.vAlign === 'top' ? null : cell.vAlign
    if (cellTouched.has('wrap')) patch.noWrap = cell.wrap ? null : true
    if (cellTouched.has('fitText')) patch.fitText = cell.fitText ? true : null
    const marginKeys: Array<keyof CellValue> = [
      'sameMargins',
      'marginTopCm',
      'marginRightCm',
      'marginBottomCm',
      'marginLeftCm',
    ]
    if (marginKeys.some((key) => cellTouched.has(key))) {
      const twips = (cm: number) => Math.max(0, Math.round(cm * TWIPS_PER_CM))
      patch.cellMar = cell.sameMargins
        ? null
        : {
            top: twips(cell.marginTopCm),
            right: twips(cell.marginRightCm),
            bottom: twips(cell.marginBottomCm),
            left: twips(cell.marginLeftCm),
          }
    }
    if (Object.keys(patch).length) run(setCellProps(patch))
    setCellTouched(new Set())
  }
  const commitAlt = () => {
    if (!altTouched) return
    run(
      updateSelectedTableAttrs({
        tblCaption: altTitle.trim() || null,
        tblDescription: altDescription.trim() || null,
        tblAltEdited: true,
      }),
    )
    setAltTouched(false)
  }
  const reload = () => {
    const next = selectionOf(editor, sectionContentWidthPx)
    setSel(next)
    setRow(next?.row ?? null)
    setColumnWidth(next?.width ?? null)
    setCell(next?.cell ?? null)
  }
  const step = (dRow: number, dCol: number) => {
    if (editor.isEditable) {
      commitRow()
      commitColumn()
      commitCell()
    }
    run(moveCaretInTable(dRow, dCol))
    reload()
  }
  const ok = () => {
    if (editableOrToast(editor)) {
      applyTableValue(editor, tableValue, tableTouched, sectionContentWidthPx)
      commitRow()
      commitColumn()
      commitCell()
      commitAlt()
    }
    onClose()
  }

  // dialog state stays in cm; the field shows and reads the preferred unit
  const cmInput = (
    value: number,
    onChange: (n: number) => void,
    opts: { min?: number; label: string; disabled?: boolean },
  ) => (
    <label>
      {opts.label}
      <span>
        <LengthInput
          value={Math.round(value * TWIPS_PER_CM)}
          min={(opts.min ?? 0) * TWIPS_PER_CM}
          max={50 * TWIPS_PER_CM}
          disabled={opts.disabled}
          ariaLabel={opts.label}
          live
          onCommit={(twips) => onChange((twips ?? 0) / TWIPS_PER_CM)}
        />
      </span>
    </label>
  )
  const widthField = (width: PreferredWidth, onChange: (next: PreferredWidth) => void) => (
    <div className="table-properties-inline table-properties-width">
      <label className="table-dialog-check">
        <Unavailable id="table-properties.preferred-width-check">
          <input type="checkbox" checked disabled readOnly />
        </Unavailable>
        {t('ribbonPreferredWidth')}
      </label>
      {width.inPct ? (
        <span className="table-properties-pct">
          <input
            type="number"
            min={1}
            max={100}
            step={1}
            aria-label={t('ribbonPreferredWidth')}
            value={width.pct}
            onChange={(e) => onChange({ ...width, pct: Number(e.target.value) || 0 })}
          />
          %
        </span>
      ) : (
        <LengthInput
          value={Math.round(width.cm * TWIPS_PER_CM)}
          min={0}
          max={50 * TWIPS_PER_CM}
          ariaLabel={t('ribbonPreferredWidth')}
          live
          onCommit={(twips) => onChange({ ...width, cm: (twips ?? 0) / TWIPS_PER_CM })}
        />
      )}
      <label>
        {t('ribbonMeasureIn')}
        <select
          value={width.inPct ? 'pct' : 'len'}
          onChange={(e) => onChange(measureWidthIn(width, e.target.value === 'pct'))}
        >
          <option value="len">{t(UNIT_NAME_KEYS[unit])}</option>
          <option value="pct">{t('ribbonZoomPercent')}</option>
        </select>
      </label>
    </div>
  )
  const rangeLabel = (a: number, b: number, one: StringKey, many: StringKey) =>
    b - a > 1 ? t(many, { a: a + 1, b }) : t(one, { n: a + 1 })

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal gs-form table-properties-dialog" role="dialog">
        <div className="table-properties-header">
          <span className="table-properties-header-icon" aria-hidden="true">
            <IconTableProperties size={22} />
          </span>
          <h2>{t('ribbonTableProperties')}</h2>
        </div>
        <div className="dlg-tabs" role="tablist">
          {TABS.map((k) => (
            <button
              key={k}
              type="button"
              role="tab"
              aria-selected={tab === k}
              className={tab === k ? 'active' : ''}
              onClick={() => setTab(k)}
            >
              {t(TAB_LABELS[k])}
            </button>
          ))}
        </div>

        {tab === 'table' && (
          <>
            <section className="table-properties-card">
              <h3>{t('ribbonAutoFit')}</h3>
              <div className="table-properties-segments">
                {TABLE_AUTO_FIT_OPTIONS.map(([mode, label]) => (
                  <label key={mode} className="table-properties-segment">
                    <input
                      type="radio"
                      name="table-autofit"
                      checked={tableValue.autoFit === mode}
                      onChange={() => updateTable('autoFit', { autoFit: mode })}
                    />
                    <span>{t(label)}</span>
                  </label>
                ))}
              </div>
            </section>
            <section className="table-properties-card">
              <h3>{t('ribbonGroupTableAlign')}</h3>
              <div className="table-properties-segments">
                {(
                  [
                    ['left', 'appAlignLeft'],
                    ['center', 'appAlignCenter'],
                    ['right', 'appAlignRight'],
                  ] as Array<[TableAlign, StringKey]>
                ).map(([mode, label]) => (
                  <label key={mode} className="table-properties-segment">
                    <input
                      type="radio"
                      name="table-align"
                      checked={tableValue.align === mode}
                      onChange={() => updateTable('align', { align: mode })}
                    />
                    <span>{t(label)}</span>
                  </label>
                ))}
              </div>
            </section>
            <section className="table-properties-card">
              <h3>{t('ribbonWrapText')}</h3>
              <div className="table-properties-segments">
                {(
                  [
                    ['none', 'appWrapInline'],
                    ['left', 'appWrapSquareLeft'],
                    ['right', 'appWrapSquareRight'],
                  ] as Array<[TableWrapMode, StringKey]>
                ).map(([mode, label]) => (
                  <label key={mode} className="table-properties-segment">
                    <input
                      type="radio"
                      name="table-wrap"
                      checked={tableValue.wrap === mode}
                      onChange={() => updateTable('wrap', { wrap: mode })}
                    />
                    <span>{t(label)}</span>
                  </label>
                ))}
              </div>
              {tableValue.wrap !== 'none' && (
                <div className="table-properties-grid">
                  {cmInput(tableValue.positionXCm, (n) => updateTable('wrap', { positionXCm: n }), {
                    min: -50,
                    label: t('ribbonHorizontalPosition'),
                  })}
                  {cmInput(tableValue.positionYCm, (n) => updateTable('wrap', { positionYCm: n }), {
                    min: -50,
                    label: t('ribbonVerticalPosition'),
                  })}
                  {cmInput(tableValue.distanceCm, (n) => updateTable('wrap', { distanceCm: n }), {
                    label: t('ribbonDistanceFromText'),
                  })}
                </div>
              )}
            </section>
            <section className="table-properties-card">
              <h3>{t('ribbonCellMargins')}</h3>
              <div className="table-properties-grid">
                {cmInput(
                  tableValue.marginTopCm,
                  (n) => updateTable('margins', { marginTopCm: n }),
                  {
                    label: t('ribbonMarginTop'),
                  },
                )}
                {cmInput(
                  tableValue.marginBottomCm,
                  (n) => updateTable('margins', { marginBottomCm: n }),
                  { label: t('ribbonMarginBottom') },
                )}
                {cmInput(
                  tableValue.marginLeftCm,
                  (n) => updateTable('margins', { marginLeftCm: n }),
                  {
                    label: t('ribbonMarginLeft'),
                  },
                )}
                {cmInput(
                  tableValue.marginRightCm,
                  (n) => updateTable('margins', { marginRightCm: n }),
                  { label: t('ribbonMarginRight') },
                )}
              </div>
            </section>
          </>
        )}

        {tab === 'row' && sel && row && (
          <section className="table-properties-card">
            <h3>{rangeLabel(sel.top, sel.bottom, 'ribbonRowN', 'ribbonRowsRange')}</h3>
            <div className="table-properties-inline">
              <label className="table-dialog-check">
                <input
                  type="checkbox"
                  checked={row.specifyHeight}
                  onChange={(e) => updateRow({ specifyHeight: e.target.checked })}
                />
                {t('ribbonRowSpecifyHeight')}
              </label>
              {cmInput(row.heightCm, (n) => updateRow({ heightCm: n, specifyHeight: true }), {
                label: t('ribbonRowHeight'),
                disabled: !row.specifyHeight,
              })}
              <label>
                {t('ribbonRowHeightIs')}
                <select
                  value={row.rule}
                  disabled={!row.specifyHeight}
                  onChange={(e) => updateRow({ rule: e.target.value as RowValue['rule'] })}
                >
                  <option value="atLeast">{t('ribbonHeightAtLeast')}</option>
                  <option value="exact">{t('ribbonHeightExactly')}</option>
                </select>
              </label>
            </div>
            <label className="table-dialog-check">
              <input
                type="checkbox"
                checked={row.allowBreak}
                onChange={(e) => updateRow({ allowBreak: e.target.checked })}
              />
              {t('ribbonAllowRowBreak')}
            </label>
            <label className={`table-dialog-check${sel.top === 0 ? '' : ' disabled'}`}>
              <input
                type="checkbox"
                disabled={sel.top !== 0}
                checked={row.repeatHeader}
                onChange={(e) => updateRow({ repeatHeader: e.target.checked })}
              />
              {t('ribbonRepeatHeaderRows')}
            </label>
            <div className="table-properties-nav">
              <button type="button" disabled={sel.top === 0} onClick={() => step(-1, 0)}>
                {t('ribbonPreviousRow')}
              </button>
              <button type="button" disabled={sel.bottom >= sel.rows} onClick={() => step(1, 0)}>
                {t('ribbonNextRow')}
              </button>
            </div>
          </section>
        )}

        {tab === 'column' && sel && columnWidth && (
          <section className="table-properties-card">
            <h3>{rangeLabel(sel.left, sel.right, 'ribbonColumnN', 'ribbonColumnsRange')}</h3>
            {widthField(columnWidth, (next) => {
              setColumnWidth(next)
              setColumnTouched(true)
            })}
            <div className="table-properties-nav">
              <button type="button" disabled={sel.left === 0} onClick={() => step(0, -1)}>
                {t('ribbonPreviousColumn')}
              </button>
              <button type="button" disabled={sel.right >= sel.cols} onClick={() => step(0, 1)}>
                {t('ribbonNextColumn')}
              </button>
            </div>
          </section>
        )}

        {tab === 'cell' && sel && cell && (
          <>
            <section className="table-properties-card">
              <h3>{t('ribbonGroupCellSize')}</h3>
              {widthField(cell.width, (next) => updateCell({ width: next }))}
            </section>
            <section className="table-properties-card">
              <h3>{t('ribbonCellVAlign')}</h3>
              <div className="table-properties-segments">
                {(
                  [
                    ['top', 'ribbonAlignTop'],
                    ['center', 'ribbonAlignMiddle'],
                    ['bottom', 'ribbonAlignBottom'],
                  ] as Array<[CellVAlign, StringKey]>
                ).map(([mode, label]) => (
                  <label key={mode} className="table-properties-segment">
                    <input
                      type="radio"
                      name="cell-valign"
                      checked={cell.vAlign === mode}
                      onChange={() => updateCell({ vAlign: mode })}
                    />
                    <span>{t(label)}</span>
                  </label>
                ))}
              </div>
            </section>
            <div className="table-properties-nav">
              <button
                type="button"
                onClick={() => setCellOptions({ cell, touched: new Set(cellTouched) })}
              >
                {t('ribbonPbOptions')}
              </button>
            </div>
          </>
        )}

        {tab === 'alt' && (
          <section className="table-properties-card table-alt-text">
            <label>
              {t('ribbonAltTitle')}
              <input
                type="text"
                value={altTitle}
                onChange={(e) => {
                  setAltTitle(e.target.value)
                  setAltTouched(true)
                }}
              />
            </label>
            <label>
              {t('ribbonAltDescription')}
              <textarea
                rows={4}
                value={altDescription}
                onChange={(e) => {
                  setAltDescription(e.target.value)
                  setAltTouched(true)
                }}
              />
            </label>
            <p className="modal-desc">{t('ribbonAltTextHint')}</p>
          </section>
        )}

        <div className="modal-actions">
          <button type="button" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button type="button" className="primary" onClick={ok}>
            {t('ribbonOk')}
          </button>
        </div>
      </div>
      {cellOptions && cell && (
        <div
          className="modal-backdrop"
          onKeyDown={(e) => {
            if (e.key !== 'Escape') return
            e.stopPropagation()
            cancelCellOptions()
          }}
        >
          <div
            className="modal gs-form table-properties-dialog table-cell-options"
            role="dialog"
            aria-label={t('ribbonCellOptionsTitle')}
          >
            <h2>{t('ribbonCellOptionsTitle')}</h2>
            <section className="table-properties-card">
              <h3>{t('ribbonCellMargins')}</h3>
              <label className="table-dialog-check">
                <input
                  type="checkbox"
                  autoFocus
                  checked={cell.sameMargins}
                  onChange={(e) => updateCell({ sameMargins: e.target.checked })}
                />
                {t('ribbonSameAsTable')}
              </label>
              <div className="table-properties-grid">
                {cmInput(
                  cell.marginTopCm,
                  (n) => updateCell({ marginTopCm: n, sameMargins: false }),
                  {
                    label: t('ribbonMarginTop'),
                    disabled: cell.sameMargins,
                  },
                )}
                {cmInput(
                  cell.marginBottomCm,
                  (n) => updateCell({ marginBottomCm: n, sameMargins: false }),
                  { label: t('ribbonMarginBottom'), disabled: cell.sameMargins },
                )}
                {cmInput(
                  cell.marginLeftCm,
                  (n) => updateCell({ marginLeftCm: n, sameMargins: false }),
                  {
                    label: t('ribbonMarginLeft'),
                    disabled: cell.sameMargins,
                  },
                )}
                {cmInput(
                  cell.marginRightCm,
                  (n) => updateCell({ marginRightCm: n, sameMargins: false }),
                  { label: t('ribbonMarginRight'), disabled: cell.sameMargins },
                )}
              </div>
            </section>
            <section className="table-properties-card">
              <h3>{t('ribbonGroupHfOptions')}</h3>
              <label className="table-dialog-check">
                <input
                  type="checkbox"
                  checked={cell.wrap}
                  onChange={(e) => updateCell({ wrap: e.target.checked })}
                />
                {t('ribbonCellWrapText')}
              </label>
              <label className="table-dialog-check">
                <input
                  type="checkbox"
                  name="cell-fit-text"
                  checked={cell.fitText}
                  onChange={(e) => updateCell({ fitText: e.target.checked })}
                />
                {t('ribbonCellFitText')}
              </label>
            </section>
            <div className="modal-actions">
              <button type="button" onClick={cancelCellOptions}>
                {t('ribbonCancel')}
              </button>
              <button type="button" className="primary" onClick={() => setCellOptions(null)}>
                {t('ribbonOk')}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

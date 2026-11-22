import { useEffect, useState } from 'react'
import type { CSSProperties } from 'react'
import type { Editor } from '@tiptap/core'
import { ShapePreview, WORDART_PRESETS, wordArtStrokePx } from '@genoffice/ui'
import { scatterXValues, type ChartDisplay, type NewChart } from '@genoffice/docx-engine'
import {
  CHART_TYPES,
  CHART_TYPE_LABELS,
  SINGLE_SERIES_TYPES,
  STOCK_MAX_SERIES,
  STOCK_MIN_SERIES,
  chartKindOf,
  type ChartType,
} from './chart-types'
import { EquationGallery } from './EquationModal'
import { COVER_PRESETS, coverPageRange, type CoverPreset } from '../editor/cover-pages'
import { useI18n, type StringKey } from '../i18n/locale'
import { isPlanned, type ActionCommandId } from '../../shared/commands'
import { NONE } from '../commands/menu-commands'
import { RbCommand, useCommandGreyed, useRibbonCommand, useRunCommand } from './ribbon-command'
import { useModalKeys } from './modal-keys'
export { LinkInsertModal } from './LinkDialog'
import {
  IconBook,
  IconCaret,
  IconComment,
  IconDoc,
  IconEquation,
  IconFooter,
  IconHeader,
  IconLink,
  IconPageBreak,
  IconPageNumber,
  IconChart,
  IconPicture,
  IconRefresh,
  IconShapes,
  IconSymbol,
  IconTable,
  IconTextBox,
  IconWordArt,
  IconIconSet,
  IconDateTime,
  IconNumberSign,
} from './icons'

/** icon size for the big icon-over-label ribbon buttons (slides ribbon parity) */
import {
  BIG,
  InsertTabProps,
  DOC_SHAPE_GROUPS,
  insertTableAt,
  insertTopLevelBlockAtSelection,
  MAX_TABLE_COLS,
  MAX_TABLE_ROWS,
  toggleDropdown,
} from './ribbon-tabs'

const SYMBOLS = [
  '—',
  '…',
  '·',
  '§',
  '¶',
  '†',
  '©',
  '®',
  '™',
  '°',
  '±',
  '×',
  '÷',
  '≤',
  '≥',
  '≠',
  '≈',
  '∞',
  '√',
  'µ',
  '€',
  '£',
  '¥',
  '¢',
  '←',
  '→',
  '↑',
  '↓',
  '⇐',
  '⇒',
  '★',
  '☆',
  '✓',
  '✗',
  '●',
  '○',
  '■',
  '□',
  '◆',
  '◇',
  '①',
  '②',
  '③',
  '④',
  '⑤',
  '⑥',
  '⑦',
  '⑧',
  '⑨',
  '⑩',
]

/** Word's Insert Chart: pick a type + data grid; after inserting, the data table below the chart is still editable in place */
export function ChartInsertModal({
  editor,
  initialType = 'bar',
  pick = 0,
  onClose,
}: {
  editor: Editor
  /** the type picked in Insert > Chart */
  initialType?: ChartType
  /** counts the menu picks: a pick made while the dialog is open moves it to that type */
  pick?: number
  onClose: () => void
}) {
  const { t } = useI18n()
  const [type, setType] = useState<ChartType>(initialType)
  useEffect(() => setType(initialType), [initialType, pick])
  const [title, setTitle] = useState(() => t('ribbonChartTitlePh'))
  const [categories, setCategories] = useState(() =>
    [1, 2, 3].map((n) => t('ribbonCategoryN', { n })),
  )
  const [series, setSeries] = useState(() => [
    { name: t('ribbonSeriesN', { n: 1 }), values: ['4', '6', '5'] },
    { name: t('ribbonSeriesN', { n: 2 }), values: ['2', '5', '7'] },
  ])
  // the grid starts on the series the type reads (a stock chart's three, the one series of a
  // funnel), unless it was already edited
  const [edited, setEdited] = useState(false)
  const stock = type === 'stock'
  const seed = stock
    ? 'stock'
    : type === 'waterfall'
      ? 'steps'
      : SINGLE_SERIES_TYPES.has(type)
        ? 'one'
        : 'two'
  useEffect(() => {
    if (edited) return
    const first = { name: t('ribbonSeriesN', { n: 1 }), values: ['4', '6', '5'] }
    setSeries(
      seed === 'stock'
        ? [
            { name: t('ribbonChartHigh'), values: ['55', '57', '57'] },
            { name: t('ribbonChartLow'), values: ['11', '12', '13'] },
            { name: t('ribbonChartClose'), values: ['32', '35', '34'] },
          ]
        : seed === 'steps'
          ? [{ ...first, values: ['100', '-40', '60'] }]
          : seed === 'one'
            ? [first]
            : [first, { name: t('ribbonSeriesN', { n: 2 }), values: ['2', '5', '7'] }],
    )
    // the seed follows the type only; `t` is not referentially stable
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [seed, edited])

  const setCat = (i: number, v: string) =>
    setCategories((prev) => prev.map((c, j) => (j === i ? v : c)))
  const setSerName = (i: number, v: string) => {
    setEdited(true)
    setSeries((prev) => prev.map((s, j) => (j === i ? { ...s, name: v } : s)))
  }
  const setSerVal = (i: number, c: number, v: string) => {
    setEdited(true)
    setSeries((prev) =>
      prev.map((s, j) =>
        j === i ? { ...s, values: s.values.map((x, k) => (k === c ? v : x)) } : s,
      ),
    )
  }
  const addCategory = () => {
    setEdited(true)
    setCategories((prev) => [...prev, t('ribbonCategoryN', { n: prev.length + 1 })])
    setSeries((prev) => prev.map((s) => ({ ...s, values: [...s.values, ''] })))
  }
  const addSeries = () => {
    setEdited(true)
    setSeries((prev) =>
      // a stock chart's fourth series is the open, and it is read first
      stock
        ? [{ name: t('ribbonChartOpen'), values: categories.map(() => '') }, ...prev]
        : [
            ...prev,
            { name: t('ribbonSeriesN', { n: prev.length + 1 }), values: categories.map(() => '') },
          ],
    )
  }

  const insert = () => {
    const spec: NewChart = {
      ...chartKindOf(type),
      title: title.trim() || t('ribbonChartTitlePh'),
      categories,
      // a type that draws one series keeps only that one, whatever the grid held before it was picked
      series: (SINGLE_SERIES_TYPES.has(type) ? series.slice(0, 1) : series).map((s) => ({
        name: s.name,
        values: s.values.map((v) => {
          const n = Number(v.trim().replace(/,/g, ''))
          return v.trim() !== '' && Number.isFinite(n) ? n : null
        }),
      })),
    }
    const display: ChartDisplay = {
      partPath: '',
      kind: spec.kind,
      ...(spec.layout ? { layout: spec.layout } : {}),
      ...(spec.horizontal ? { horizontal: true } : {}),
      title: spec.title,
      categories: spec.categories,
      // a scatter point sits at the x its category names, as the saved chart will read back
      series:
        spec.kind === 'scatter'
          ? spec.series.map((s) => ({ ...s, xValues: scatterXValues(spec.categories) }))
          : spec.series,
    }
    const inserted = insertTopLevelBlockAtSelection(editor, {
      type: 'docProtected',
      attrs: {
        docxIndex: null,
        blockType: 'chart',
        label: t('ribbonChart'),
        genChart: spec,
        chartDisplay: display,
      },
    })
    if (inserted) onClose()
  }

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal modal-chart">
        <h2>{t('ribbonChartInsertTitle')}</h2>
        <div className="modal-row chart-type-options">
          {CHART_TYPES.map((value) => (
            <button
              key={value}
              data-chart-type={value}
              className={type === value ? 'btn-primary' : ''}
              // command-exempt: picks the type inside the chart dialog
              onClick={() => setType(value)}
            >
              {t(CHART_TYPE_LABELS[value])}
            </button>
          ))}
        </div>
        <div className="modal-row">
          <input
            value={title}
            placeholder={t('ribbonChartTitlePh')}
            onChange={(e) => setTitle(e.target.value)}
          />
        </div>
        <table className="chart-data-grid">
          <thead>
            <tr>
              <th />
              {categories.map((c, i) => (
                <th key={i}>
                  <input value={c} onChange={(e) => setCat(i, e.target.value)} />
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {series.map((s, i) => (
              <tr key={i}>
                <th>
                  <input value={s.name} onChange={(e) => setSerName(i, e.target.value)} />
                </th>
                {s.values.map((v, c) => (
                  <td key={c}>
                    <input
                      value={v}
                      inputMode="decimal"
                      onChange={(e) => setSerVal(i, c, e.target.value)}
                    />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
        <div className="modal-row">
          <button onClick={addCategory}>{t('ribbonChartAddCategory')}</button>
          {/* command-exempt: adds a series to the dialog's own data grid */}
          <button
            onClick={addSeries}
            // a type that draws one series takes no second one: it would be saved and never shown
            disabled={
              type === 'pie' ||
              SINGLE_SERIES_TYPES.has(type) ||
              (stock && series.length >= STOCK_MAX_SERIES)
            }
          >
            {t('ribbonChartAddSeries')}
          </button>
        </div>
        <div className="modal-actions">
          <button
            className="btn-primary"
            // command-exempt: confirms the chart dialog; the command opened it
            onClick={insert}
            disabled={
              stock && (series.length < STOCK_MIN_SERIES || series.length > STOCK_MAX_SERIES)
            }
          >
            {t('ribbonInsert')}
          </button>
          <button onClick={onClose}>{t('ribbonCancel')}</button>
        </div>
      </div>
    </div>
  )
}

/** Word's Insert Table dialog: explicit row/column counts beyond the hover grid's reach */
export function TableInsertModal({
  editor,
  onClose,
  onInserted,
}: {
  editor: Editor
  onClose: () => void
  onInserted?: () => void
}) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [cols, setCols] = useState(5)
  const [rows, setRows] = useState(2)

  const insert = () => {
    if (!editor.isEditable) return
    insertTableAt(editor, rows, cols)
    onInserted?.()
    onClose()
  }

  const countInput = (label: string, value: number, max: number, set: (v: number) => void) => (
    <label>
      {label}
      <input
        type="number"
        min={1}
        max={max}
        value={value}
        onChange={(e) => {
          const v = Math.round(Number(e.target.value))
          set(Number.isFinite(v) ? Math.min(max, Math.max(1, v)) : 1)
        }}
        onKeyDown={(e) => e.key === 'Enter' && insert()}
      />
    </label>
  )

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal">
        <h2>{t('ribbonTableInsertDialog')}</h2>
        <div className="font-dialog-row">
          {countInput(t('ribbonTableColsLabel'), cols, MAX_TABLE_COLS, setCols)}
          {countInput(t('ribbonTableRowsLabel'), rows, MAX_TABLE_ROWS, setRows)}
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" onClick={insert}>
            {t('ribbonOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

/** Mini cover-page preview: scaled render of the preset data (twips ≈ /130 px, font size ≈ /14 px) */
function CoverThumb({ preset }: { preset: CoverPreset }) {
  return (
    <div className="cover-thumb-page">
      {preset.paras.map((p, i) => {
        const style: CSSProperties = {
          textAlign: p.align ?? 'left',
          marginTop: p.spaceBefore ? p.spaceBefore / 130 : 0,
          marginBottom: p.spaceAfter ? p.spaceAfter / 130 : 0,
          marginLeft: p.indentLeft ? p.indentLeft / 130 : 0,
          marginRight: p.indentRight ? p.indentRight / 130 : 0,
        }
        if (p.shadingFill) style.backgroundColor = `#${p.shadingFill}`
        if (p.color) style.color = `#${p.color}`
        if (p.bold) style.fontWeight = 700
        if (p.italic) style.fontStyle = 'italic'
        if (p.sizeHalfPoints) style.fontSize = Math.max(3, p.sizeHalfPoints / 14)
        if (p.borders) {
          const line = '0.5px solid #666'
          if (p.borders.includes('t')) style.borderTop = line
          if (p.borders.includes('b')) style.borderBottom = line
          if (p.borders.includes('l')) style.borderLeft = line
          if (p.borders.includes('r')) style.borderRight = line
        }
        return (
          <div key={i} className="cover-thumb-para" style={style}>
            {p.text ?? '\u00a0'}
          </div>
        )
      })}
    </div>
  )
}

export function InsertTab({
  editor,
  hasDoc,
  dropdown,
  setDropdown,
  header,
  footer,
  isProtected,
  commentsAllowed,
}: InsertTabProps) {
  const { t } = useI18n()
  const [grid, setGrid] = useState<{ r: number; c: number }>({ r: 0, c: 0 })
  const runCommand = useRunCommand()
  // commenting outlives the comments-only restriction, which the protection terms below answer
  const commentOff = useCommandGreyed('insert-comment', { editable: true })
  const dropCapOff = useCommandGreyed('drop-cap') || !hasDoc
  const convertText = useRibbonCommand('convert-text-to-table')
  const textFromFile = useRibbonCommand('insert-file')
  /** a pick from a gallery: run the command, close the gallery */
  const pick = (id: ActionCommandId, payload?: string) => {
    runCommand(id, payload)
    setDropdown(() => null)
  }

  return (
    <>
      {/* Word: Pages (cover page / blank page / page break) sits leftmost */}
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-col" data-wide="2">
            <div className="rb-split-wrap">
              <button
                className="rb-small"
                disabled={!hasDoc}
                data-tip={t('ribbonCoverPageTip')}
                onClick={() => toggleDropdown(setDropdown, 'cover')}
              >
                <IconDoc size={14} />
                <span className="rb-small-label">{t('ribbonCoverPage')}</span>
                <IconCaret />
              </button>
              {dropdown === 'cover' && (
                <div data-rb-panel="" className="cover-gallery">
                  {COVER_PRESETS.map((preset) => (
                    <button
                      key={preset.id}
                      className="cover-card"
                      data-tip={t('ribbonInsertCoverTip', { name: preset.name })}
                      onClick={() => pick('insert-cover-page', preset.id)}
                    >
                      <CoverThumb preset={preset} />
                      <span className="cover-card-name">{preset.name}</span>
                    </button>
                  ))}
                  <button
                    className="cover-gallery-remove"
                    disabled={!coverPageRange(editor.state.doc)}
                    onClick={() => pick('remove-cover-page')}
                  >
                    {t('ribbonRemoveCoverPage')}
                  </button>
                </div>
              )}
            </div>
            <RbCommand cmd="insert-blank-page" look="small">
              <IconDoc size={14} />
            </RbCommand>
            <RbCommand cmd="insert-page-break" look="small">
              <IconPageBreak size={14} />
            </RbCommand>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupPages')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonTableTip')}
              onClick={() => toggleDropdown(setDropdown, 'table')}
            >
              <span className="rb-big-icon">
                <IconTable size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonTable')}</span>
            </button>
            {dropdown === 'table' && (
              <div
                data-rb-panel=""
                className="table-picker"
                onMouseLeave={() => setGrid({ r: 0, c: 0 })}
              >
                <div className="table-picker-title">
                  {grid.r > 0
                    ? t('ribbonTableSize', { r: grid.r, c: grid.c })
                    : t('ribbonTablePickSize')}
                </div>
                <div className="table-picker-grid">
                  {Array.from({ length: 8 }, (_, ri) =>
                    Array.from({ length: 10 }, (_, ci) => (
                      <button
                        key={`${ri}-${ci}`}
                        className={`table-cell ${ri < grid.r && ci < grid.c ? 'hot' : ''}`}
                        onMouseEnter={() => setGrid({ r: ri + 1, c: ci + 1 })}
                        onClick={() => pick('insert-table-grid', `${ri + 1},${ci + 1}`)}
                      />
                    )),
                  )}
                </div>
                <button
                  className="table-picker-custom"
                  onClick={() => {
                    setDropdown(() => null)
                    runCommand('insert-table')
                  }}
                >
                  {t('ribbonTableInsertDialog')}
                </button>
                <button
                  className="table-picker-custom"
                  disabled={convertText.disabled}
                  onClick={() => {
                    setDropdown(() => null)
                    convertText.run()
                  }}
                >
                  {t('ribbonConvertTextToTableTitle')}…
                </button>
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupTables')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <RbCommand cmd="insert-image" look="big">
            <IconPicture size={BIG} />
          </RbCommand>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonShapesTip')}
              onClick={() => toggleDropdown(setDropdown, 'shape')}
            >
              <span className="rb-big-icon">
                <IconShapes size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonShapes')}</span>
            </button>
            {dropdown === 'shape' && (
              <div data-rb-panel="" className="shape-palette rb-shape-gallery">
                {DOC_SHAPE_GROUPS.map((group) => (
                  <div key={group.groupKey}>
                    <div className="rb-drop-title">{t(group.groupKey as StringKey)}</div>
                    <div className="rb-shape-grid">
                      {group.shapes.map((s) => (
                        <button
                          key={s.prst}
                          className="rb-shape-cell"
                          data-tip={t(s.labelKey as StringKey)}
                          aria-label={t(s.labelKey as StringKey)}
                          onClick={() => pick('insert-shape', s.prst)}
                        >
                          <ShapePreview prst={s.prst} size={18} />
                        </button>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
          <RbCommand cmd="insert-icons" look="big">
            <IconIconSet size={BIG} />
          </RbCommand>
          <RbCommand cmd="insert-chart" look="big">
            <IconChart size={BIG} />
          </RbCommand>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupIllustrations')}</div>
      </div>

      <div className="ribbon-sep" />

      {/* Word: Link/Bookmark/Cross-reference are small vertical buttons */}
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-col" data-wide="1">
            <RbCommand cmd="insert-link" look="small">
              <IconLink size={14} />
            </RbCommand>
            <RbCommand cmd="insert-bookmark" look="small">
              <IconBook size={14} />
            </RbCommand>
            <RbCommand cmd="insert-cross-reference" look="small">
              <IconRefresh size={14} />
            </RbCommand>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupLinks')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          {/* Word: Insert → Comment starts a new comment; the pane toggle stays on Review.
              Deliberately not gated on this tab's hasDoc (= canEdit): under the comments-only
              restriction the body is read-only yet commenting stays allowed, matching the
              Review tab's New Comment. */}
          <button
            className="rb-big"
            disabled={commentOff || (isProtected && !commentsAllowed)}
            data-tip={commentOff ? t('ribbonNewCommentSelectTip') : t('ribbonNewCommentTip')}
            onClick={() => runCommand('insert-comment')}
          >
            <span className="rb-big-icon">
              <IconComment size={BIG} />
            </span>
            <span>{t('ribbonComment')}</span>
          </button>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupComments')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className={`rb-big ${header?.text ? 'active' : ''}`}
              disabled={!hasDoc}
              data-tip={t('ribbonHeaderTip')}
              onClick={() => toggleDropdown(setDropdown, 'header')}
            >
              <span className="rb-big-icon">
                <IconHeader size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonHeader')}</span>
            </button>
            {dropdown === 'header' && (
              <div data-rb-panel="" className="layout-menu">
                <button onClick={() => pick('edit-header-footer', 'header')}>
                  {t('ribbonHfEditHeader')}
                </button>
                <button onClick={() => pick('remove-header-footer', 'header')}>
                  {t('ribbonHfRemoveHeader')}
                </button>
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className={`rb-big ${footer?.text ? 'active' : ''}`}
              disabled={!hasDoc}
              data-tip={t('ribbonFooterTip')}
              onClick={() => toggleDropdown(setDropdown, 'footer')}
            >
              <span className="rb-big-icon">
                <IconFooter size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonFooter')}</span>
            </button>
            {dropdown === 'footer' && (
              <div data-rb-panel="" className="layout-menu">
                <button onClick={() => pick('edit-header-footer', 'footer')}>
                  {t('ribbonHfEditFooter')}
                </button>
                <button onClick={() => pick('remove-header-footer', 'footer')}>
                  {t('ribbonHfRemoveFooter')}
                </button>
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className={`rb-big ${footer?.pageNumber ? 'active' : ''}`}
              disabled={!hasDoc}
              data-tip={t('ribbonPageNumber')}
              onClick={() => toggleDropdown(setDropdown, 'pagenum')}
            >
              <span className="rb-big-icon">
                <IconPageNumber size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonPageNumber')}</span>
            </button>
            {dropdown === 'pagenum' && (
              <div data-rb-panel="" className="layout-menu">
                {(
                  [
                    [t('ribbonPnTopLeft'), 'header', 'left'],
                    [t('ribbonPnTopCenter'), 'header', 'center'],
                    [t('ribbonPnTopRight'), 'header', 'right'],
                    [t('ribbonPnBottomLeft'), 'footer', 'left'],
                    [t('ribbonPnBottomCenter'), 'footer', 'center'],
                    [t('ribbonPnBottomRight'), 'footer', 'right'],
                  ] as Array<[string, 'header' | 'footer', 'left' | 'center' | 'right']>
                ).map(([label, kind, align]) => (
                  <button key={label} onClick={() => pick('page-number', `${kind},${align}`)}>
                    {label}
                  </button>
                ))}
                <button onClick={() => pick('page-number-format')}>{t('ribbonPnFormat')}</button>
                <button onClick={() => pick('page-number-remove')}>{t('ribbonPnRemove')}</button>
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupHeaderFooter')}</div>
      </div>

      <div className="ribbon-sep" />

      {/* Text group: text box + WordArt + drop cap */}
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonTextBoxTip')}
              onClick={() => toggleDropdown(setDropdown, 'textBox')}
            >
              <span className="rb-big-icon">
                <IconTextBox size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonTextBox')}</span>
            </button>
            {dropdown === 'textBox' && (
              <div data-rb-panel="" className="layout-menu rb-textbox-menu">
                <button onClick={() => pick('insert-text-box', 'draw')}>
                  {t('ribbonTextBoxDraw')}
                </button>
                <button onClick={() => pick('insert-vertical-text-box')}>
                  {t('ribbonTextBoxDrawVertical')}
                </button>
                <button onClick={() => pick('insert-text-box', 'simple')}>
                  {t('ribbonTextBoxSimple')}
                </button>
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonWordArtTip')}
              onClick={() => toggleDropdown(setDropdown, 'wordArt')}
            >
              <span className="rb-big-icon">
                <IconWordArt size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonWordArt')}</span>
            </button>
            {dropdown === 'wordArt' && (
              <div data-rb-panel="" className="wordart-palette">
                {WORDART_PRESETS.map((p) => (
                  <button
                    key={p.id}
                    className="wordart-cell"
                    data-tip={t(p.nameKey as StringKey)}
                    style={{
                      color: p.fill,
                      WebkitTextStroke: p.outline
                        ? `${wordArtStrokePx(p.outline.widthEmu)}px ${p.outline.color}`
                        : undefined,
                      fontWeight: p.bold ? 800 : 400,
                      fontStyle: p.italic ? 'italic' : undefined,
                    }}
                    onClick={() => pick('insert-wordart', p.id)}
                  >
                    {t('ribbonWordArtPreviewChar')}
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={dropCapOff}
              data-tip={t('ribbonDropCap')}
              onClick={() => toggleDropdown(setDropdown, 'dropCap')}
            >
              <span
                className="rb-big-icon"
                style={{ fontSize: 20, lineHeight: 1, fontWeight: 'bold' }}
              >
                A̲
                <IconCaret />
              </span>
              <span>{t('ribbonDropCap')}</span>
            </button>
            {dropdown === 'dropCap' && (
              <div data-rb-panel="" className="dropcap-menu">
                {(
                  [
                    { val: null, label: t('ribbonNone'), desc: t('ribbonDropCapNoneDesc') },
                    {
                      val: 'drop',
                      label: t('ribbonDropCapDropped'),
                      desc: t('ribbonDropCapDroppedDesc'),
                    },
                    {
                      val: 'margin',
                      label: t('ribbonDropCapMargin'),
                      desc: t('ribbonDropCapMarginDesc'),
                    },
                  ] as const
                ).map((opt) => (
                  <button
                    key={opt.label}
                    className="dropcap-option"
                    data-tip={opt.desc}
                    onClick={() => pick('drop-cap', opt.val ?? NONE)}
                  >
                    <span className="dropcap-preview">{opt.val ? 'A' : '—'}</span>
                    <span>{opt.label}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="rb-col" data-collapse="2">
            <div className="rb-split-wrap">
              <button
                className="rb-small"
                disabled={!hasDoc}
                data-tip={t('ribbonFieldTip')}
                onClick={() => toggleDropdown(setDropdown, 'field')}
              >
                <span className="rb-small-glyph">{'{ }'}</span>
                <span className="rb-small-label">{t('ribbonField')}</span>
                <IconCaret />
              </button>
              {dropdown === 'field' && (
                <div data-rb-panel="" className="layout-menu">
                  {(
                    [
                      ['DATE', t('ribbonFieldDate')],
                      ['TIME', t('ribbonFieldTime')],
                      ['PAGE', t('ribbonFieldPage')],
                      ['NUMPAGES', t('ribbonFieldNumPages')],
                      ['FILENAME', t('ribbonFieldFileName')],
                    ] as const
                  ).map(([instr, label]) => (
                    <button key={instr} onClick={() => pick('insert-field', instr)}>
                      {label}
                    </button>
                  ))}
                </div>
              )}
            </div>
            <RbCommand cmd="insert-date-time" look="small">
              <IconDateTime size={14} />
            </RbCommand>
            <div className="rb-split-wrap">
              <button
                className="rb-small"
                disabled={!hasDoc}
                data-tip={t('ribbonObject')}
                onClick={() => toggleDropdown(setDropdown, 'object')}
              >
                <IconDoc size={14} />
                <span className="rb-small-label">{t('ribbonObject')}</span>
                <IconCaret />
              </button>
              {dropdown === 'object' && (
                <div data-rb-panel="" className="layout-menu">
                  {!isPlanned('insert-object') && (
                    <button onClick={() => pick('insert-object' as ActionCommandId)}>
                      {t('ribbonObject')}…
                    </button>
                  )}
                  <button disabled={textFromFile.disabled} onClick={() => pick('insert-file')}>
                    {t('ribbonTextFromFile')}…
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupText')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonEquationTip')}
              onClick={() => toggleDropdown(setDropdown, 'equation')}
            >
              <span className="rb-big-icon">
                <IconEquation size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonEquation')}</span>
            </button>
            {dropdown === 'equation' && (
              <EquationGallery
                editor={editor}
                onPick={() => setDropdown(() => null)}
                onCustom={() => {
                  setDropdown(() => null)
                  runCommand('insert-equation')
                }}
              />
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              data-tip={t('ribbonSymbolTip')}
              onClick={() => toggleDropdown(setDropdown, 'symbol')}
            >
              <span className="rb-big-icon">
                <IconSymbol size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonSymbol')}</span>
            </button>
            {dropdown === 'symbol' && (
              <div data-rb-panel="" className="symbol-palette">
                {SYMBOLS.map((s) => (
                  <button key={s} className="symbol-cell" onClick={() => pick('insert-symbol', s)}>
                    {s}
                  </button>
                ))}
              </div>
            )}
          </div>
          <RbCommand cmd="insert-number" look="big">
            <IconNumberSign size={BIG} />
          </RbCommand>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupSymbols')}</div>
      </div>
    </>
  )
}

/* ================= Design ================= */

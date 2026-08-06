import { useState } from 'react'
import type { CSSProperties } from 'react'
import type { Editor } from '@tiptap/core'
import { ShapePreview, WORDART_PRESETS, wordArtStrokePx } from '@genoffice/ui'
import type { ChartDisplay, HeaderFooter, NewChart } from '@genoffice/docx-engine'
import { EquationGallery, EquationModal } from './EquationModal'
import { COVER_PRESETS, insertCoverPage, type CoverPreset } from '../editor/cover-pages'
import { startShapeDrawMode } from '../editor/shape-draw'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'
import {
  IconBook,
  IconCaret,
  IconCheckbox,
  IconCheckboxChecked,
  IconClose,
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
} from './icons'

/** icon size for the big icon-over-label ribbon buttons (slides ribbon parity) */
import {
  BIG,
  InsertTabProps,
  DOC_SHAPE_GROUPS,
  insertBlankPageAt,
  insertImageViaDialog,
  insertPageBreakAt,
  insertShapeAt,
  insertTableAt,
  insertTextboxAt,
  insertTopLevelBlockAtSelection,
  insertWordArtAt,
  setParaAttrs,
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

/** small dropdown editor for header/footer text */
function HfEditor({
  label,
  current,
  onApply,
  onClose,
}: {
  label: string
  current: string
  onApply: (text: string) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const [text, setText] = useState(current)
  return (
    <div className="hf-menu">
      <div className="hf-menu-title">{label}</div>
      <input
        autoFocus
        value={text}
        placeholder={t('ribbonHfPlaceholder', { label })}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            onApply(text.trim())
            onClose()
          }
        }}
      />
      <div className="hf-menu-actions">
        <button
          className="btn-primary"
          onClick={() => {
            onApply(text.trim())
            onClose()
          }}
        >
          {t('ribbonOk')}
        </button>
        <button onClick={onClose}>{t('ribbonCancel')}</button>
      </div>
    </div>
  )
}

/** all bookmarks in document order: name + owning node position + preview text */
function collectBookmarks(editor: Editor): Array<{ name: string; pos: number; preview: string }> {
  const out: Array<{ name: string; pos: number; preview: string }> = []
  editor.state.doc.descendants((node, pos) => {
    const names = node.attrs?.bookmarks as string[] | null
    if (Array.isArray(names)) {
      for (const name of names) {
        out.push({ name, pos, preview: node.textContent.slice(0, 40) })
      }
    }
    return !node.isLeaf
  })
  return out
}

/** the paragraph-like node at the cursor that can carry bookmarks */
function bookmarkTargetAtCursor(editor: Editor): { pos: number; typeName: string } | null {
  const { $from } = editor.state.selection
  for (let depth = $from.depth; depth >= 1; depth--) {
    const node = $from.node(depth)
    if (['docParagraph', 'docHeading', 'docListItem'].includes(node.type.name)) {
      return { pos: $from.before(depth), typeName: node.type.name }
    }
  }
  return null
}

/** Word's Bookmark dialog: list + add (cursor's paragraph) / go to / delete */
export function BookmarkModal({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t } = useI18n()
  const [name, setName] = useState('')
  const [error, setError] = useState('')
  const [, forceRender] = useState(0)
  const bookmarks = collectBookmarks(editor)

  const addBookmark = () => {
    const trimmed = name.trim()
    if (!/^[A-Za-z一-鿿_][\w一-鿿]*$/.test(trimmed)) {
      setError(t('ribbonBookmarkNameInvalid'))
      return
    }
    if (bookmarks.some((b) => b.name === trimmed)) {
      setError(t('ribbonBookmarkExists'))
      return
    }
    const target = bookmarkTargetAtCursor(editor)
    if (!target) {
      setError(t('ribbonBookmarkNoCursor'))
      return
    }
    const node = editor.state.doc.nodeAt(target.pos)
    const existing = (node?.attrs?.bookmarks as string[] | null) ?? []
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(target.pos, undefined, {
        ...node?.attrs,
        bookmarks: [...existing, trimmed],
      }),
    )
    setName('')
    setError('')
    forceRender((n) => n + 1)
  }

  const removeBookmark = (entry: { name: string; pos: number }) => {
    const node = editor.state.doc.nodeAt(entry.pos)
    if (!node) return
    const rest = ((node.attrs?.bookmarks as string[] | null) ?? []).filter((n) => n !== entry.name)
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(entry.pos, undefined, {
        ...node.attrs,
        bookmarks: rest.length > 0 ? rest : null,
      }),
    )
    forceRender((n) => n + 1)
  }

  const jumpTo = (pos: number) => {
    const dom = editor.view.nodeDOM(pos) as HTMLElement | null
    dom?.scrollIntoView({ behavior: 'smooth', block: 'center' })
    onClose()
  }

  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <h2>{t('ribbonBookmark')}</h2>
        <div className="modal-row">
          <input
            autoFocus
            value={name}
            placeholder={t('ribbonBookmarkNamePh')}
            onChange={(e) => {
              setName(e.target.value)
              setError('')
            }}
            onKeyDown={(e) => e.key === 'Enter' && addBookmark()}
          />
          <button className="btn-primary" onClick={addBookmark}>
            {t('ribbonAdd')}
          </button>
        </div>
        {error && <div className="modal-error">{error}</div>}
        <div className="bookmark-list">
          {bookmarks.length === 0 && (
            <div className="bookmark-empty">{t('ribbonBookmarkEmpty')}</div>
          )}
          {bookmarks.map((b) => (
            <div key={`${b.name}-${b.pos}`} className="bookmark-row">
              <button className="bookmark-name" title={b.preview} onClick={() => jumpTo(b.pos)}>
                {b.name}
              </button>
              <span className="bookmark-preview">{b.preview}</span>
              <button
                className="bookmark-del"
                title={t('ribbonBookmarkDeleteTip')}
                onClick={() => removeBookmark(b)}
              >
                <IconClose size={12} />
              </button>
            </div>
          ))}
        </div>
        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonClose')}</button>
        </div>
      </div>
    </div>
  )
}

/** Word's Cross-reference dialog: pick a bookmark, insert a REF field at the cursor (displays the bookmark paragraph's text) */
export function CrossRefModal({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t } = useI18n()
  const bookmarks = collectBookmarks(editor)
  const insertRef = (b: { name: string; preview: string }) => {
    editor
      .chain()
      .focus()
      .insertContent({
        type: 'text',
        text: b.preview || b.name,
        marks: [{ type: 'refField', attrs: { name: b.name } }],
      })
      .run()
    onClose()
  }
  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <h2>{t('ribbonCrossRef')}</h2>
        {bookmarks.length === 0 ? (
          <p className="bookmark-empty">{t('ribbonCrossRefEmpty')}</p>
        ) : (
          <div className="bookmark-list">
            {bookmarks.map((b) => (
              <div key={`${b.name}-${b.pos}`} className="bookmark-row">
                <button className="bookmark-name" onClick={() => insertRef(b)}>
                  {b.name}
                </button>
                <span className="bookmark-preview">{b.preview}</span>
              </div>
            ))}
          </div>
        )}
        <div className="modal-actions">
          <button onClick={onClose}>{t('ribbonCancel')}</button>
        </div>
      </div>
    </div>
  )
}

/** Word's Insert Chart: pick a type + data grid; after inserting, the data table below the chart is still editable in place */
export function ChartInsertModal({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t } = useI18n()
  const [kind, setKind] = useState<NewChart['kind']>('bar')
  const [title, setTitle] = useState(() => t('ribbonChartTitlePh'))
  const [categories, setCategories] = useState(() =>
    [1, 2, 3].map((n) => t('ribbonCategoryN', { n })),
  )
  const [series, setSeries] = useState(() => [
    { name: t('ribbonSeriesN', { n: 1 }), values: ['4', '6', '5'] },
    { name: t('ribbonSeriesN', { n: 2 }), values: ['2', '5', '7'] },
  ])

  const setCat = (i: number, v: string) =>
    setCategories((prev) => prev.map((c, j) => (j === i ? v : c)))
  const setSerName = (i: number, v: string) =>
    setSeries((prev) => prev.map((s, j) => (j === i ? { ...s, name: v } : s)))
  const setSerVal = (i: number, c: number, v: string) =>
    setSeries((prev) =>
      prev.map((s, j) =>
        j === i ? { ...s, values: s.values.map((x, k) => (k === c ? v : x)) } : s,
      ),
    )
  const addCategory = () => {
    setCategories((prev) => [...prev, t('ribbonCategoryN', { n: prev.length + 1 })])
    setSeries((prev) => prev.map((s) => ({ ...s, values: [...s.values, ''] })))
  }
  const addSeries = () =>
    setSeries((prev) => [
      ...prev,
      { name: t('ribbonSeriesN', { n: prev.length + 1 }), values: categories.map(() => '') },
    ])

  const insert = () => {
    const spec: NewChart = {
      kind,
      title: title.trim() || t('ribbonChartTitlePh'),
      categories,
      series: series.map((s) => ({
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
      title: spec.title,
      categories: spec.categories,
      series: spec.series,
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
        <div className="modal-row">
          {(
            [
              ['bar', t('ribbonChartBar')],
              ['line', t('ribbonChartLine')],
              ['pie', t('ribbonChartPie')],
            ] as const
          ).map(([value, label]) => (
            <button
              key={value}
              className={kind === value ? 'btn-primary' : ''}
              onClick={() => setKind(value)}
            >
              {label}
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
          <button onClick={addSeries} disabled={kind === 'pie'}>
            {t('ribbonChartAddSeries')}
          </button>
        </div>
        <div className="modal-actions">
          <button className="btn-primary" onClick={insert}>
            {t('ribbonInsert')}
          </button>
          <button onClick={onClose}>{t('ribbonCancel')}</button>
        </div>
      </div>
    </div>
  )
}

/** Insert Hyperlink dialog, shared by the ribbon and the native application menu */
export function LinkInsertModal({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const [linkText, setLinkText] = useState('')
  const [linkUrl, setLinkUrl] = useState('')

  const insertLink = () => {
    const href = linkUrl.trim()
    const text = linkText.trim() || href
    if (!href || !editor.isEditable) return
    editor
      .chain()
      .focus()
      .insertContent({ type: 'text', text, marks: [{ type: 'link', attrs: { href, rId: null } }] })
      .run()
    onClose()
  }

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal">
        <h2>{t('ribbonLinkInsertTitle')}</h2>
        <label>
          {t('ribbonLinkText')}
          <input
            value={linkText}
            onChange={(e) => setLinkText(e.target.value)}
            placeholder={t('ribbonLinkTextPh')}
          />
        </label>
        <label>
          {t('ribbonLinkAddress')}
          <input
            value={linkUrl}
            onChange={(e) => setLinkUrl(e.target.value)}
            placeholder="https://…"
            onKeyDown={(e) => e.key === 'Enter' && insertLink()}
          />
        </label>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('ribbonCancel')}
          </button>
          <button className="btn-primary" disabled={!linkUrl.trim()} onClick={insertLink}>
            {t('ribbonInsert')}
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
  onHeader,
  footer,
  onFooter,
  onPageNumFormat,
  onInsertField,
  titlePg,
  onTitlePg,
  evenOddHf,
  onEvenOddHf,
  commentCount,
  onShowComments,
}: InsertTabProps) {
  const { t } = useI18n()
  const [grid, setGrid] = useState<{ r: number; c: number }>({ r: 0, c: 0 })
  const [linkOpen, setLinkOpen] = useState(false)
  const [equationOpen, setEquationOpen] = useState(false)
  const [bookmarkOpen, setBookmarkOpen] = useState(false)
  const [crossRefOpen, setCrossRefOpen] = useState(false)
  const [chartOpen, setChartOpen] = useState(false)

  const insertTable = (rows: number, cols: number) => {
    insertTableAt(editor, rows, cols)
    setDropdown(() => null)
  }

  return (
    <>
      {/* Word: Pages (cover page / blank page / page break) sits leftmost */}
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              title={t('ribbonCoverPageTip')}
              onClick={() => toggleDropdown(setDropdown, 'cover')}
            >
              <span className="rb-big-icon">
                <IconDoc size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonCoverPage')}</span>
            </button>
            {dropdown === 'cover' && (
              <div className="cover-gallery">
                {COVER_PRESETS.map((preset) => (
                  <button
                    key={preset.id}
                    className="cover-card"
                    title={t('ribbonInsertCoverTip', { name: preset.name })}
                    onClick={() => {
                      insertCoverPage(editor, preset)
                      setDropdown(() => null)
                    }}
                  >
                    <CoverThumb preset={preset} />
                    <span className="cover-card-name">{preset.name}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonBlankPageTip')}
            onClick={() => insertBlankPageAt(editor)}
          >
            <span className="rb-big-icon">
              <IconDoc size={BIG} />
            </span>
            <span>{t('ribbonBlankPage')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonPageBreakTip')}
            onClick={() => insertPageBreakAt(editor)}
          >
            <span className="rb-big-icon">
              <IconPageBreak size={BIG} />
            </span>
            <span>{t('ribbonPageBreak')}</span>
          </button>
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
              title={t('ribbonTableTip')}
              onClick={() => toggleDropdown(setDropdown, 'table')}
            >
              <span className="rb-big-icon">
                <IconTable size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonTable')}</span>
            </button>
            {dropdown === 'table' && (
              <div className="table-picker" onMouseLeave={() => setGrid({ r: 0, c: 0 })}>
                <div className="table-picker-title">
                  {grid.r > 0
                    ? t('ribbonTableSize', { r: grid.r, c: grid.c })
                    : t('ribbonTablePickSize')}
                </div>
                <div className="table-picker-grid">
                  {Array.from({ length: 6 }, (_, ri) =>
                    Array.from({ length: 8 }, (_, ci) => (
                      <button
                        key={`${ri}-${ci}`}
                        className={`table-cell ${ri < grid.r && ci < grid.c ? 'hot' : ''}`}
                        onMouseEnter={() => setGrid({ r: ri + 1, c: ci + 1 })}
                        onClick={() => insertTable(ri + 1, ci + 1)}
                      />
                    )),
                  )}
                </div>
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupTables')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonPictureTip')}
            onClick={() => void insertImageViaDialog(editor)}
          >
            <span className="rb-big-icon">
              <IconPicture size={BIG} />
            </span>
            <span>{t('ribbonPicture')}</span>
          </button>
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonChartTip')}
            onClick={() => setChartOpen(true)}
          >
            <span className="rb-big-icon">
              <IconChart size={BIG} />
            </span>
            <span>{t('ribbonChart')}</span>
          </button>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              title={t('ribbonShapesTip')}
              onClick={() => toggleDropdown(setDropdown, 'shape')}
            >
              <span className="rb-big-icon">
                <IconShapes size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonShapes')}</span>
            </button>
            {dropdown === 'shape' && (
              <div className="shape-palette rb-shape-gallery">
                {DOC_SHAPE_GROUPS.map((group) => (
                  <div key={group.groupKey}>
                    <div className="rb-drop-title">{t(group.groupKey as StringKey)}</div>
                    <div className="rb-shape-grid">
                      {group.shapes.map((s) => (
                        <button
                          key={s.prst}
                          className="rb-shape-cell"
                          title={t(s.labelKey as StringKey)}
                          onClick={() => {
                            // Word parity: arm crosshair draw mode (click = default 1in
                            // square, drag = custom size, Shift = square, Esc = cancel)
                            startShapeDrawMode(editor, s.prst, (opts) =>
                              insertShapeAt(editor, s.prst, opts),
                            )
                            setDropdown(() => null)
                          }}
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
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupIllustrations')}</div>
      </div>

      <div className="ribbon-sep" />

      {/* Text group: text box + WordArt + drop cap */}
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={t('ribbonTextBoxTip')}
            onClick={() => insertTextboxAt(editor)}
          >
            <span className="rb-big-icon">
              <IconTextBox size={BIG} />
            </span>
            <span>{t('ribbonTextBox')}</span>
          </button>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              title={t('ribbonWordArtTip')}
              onClick={() => toggleDropdown(setDropdown, 'wordArt')}
            >
              <span className="rb-big-icon">
                <IconWordArt size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonWordArt')}</span>
            </button>
            {dropdown === 'wordArt' && (
              <div className="wordart-palette">
                {WORDART_PRESETS.map((p) => (
                  <button
                    key={p.id}
                    className="wordart-cell"
                    title={t(p.nameKey as StringKey)}
                    style={{
                      color: p.fill,
                      WebkitTextStroke: p.outline
                        ? `${wordArtStrokePx(p.outline.widthEmu)}px ${p.outline.color}`
                        : undefined,
                      fontWeight: p.bold ? 800 : 400,
                      fontStyle: p.italic ? 'italic' : undefined,
                    }}
                    onClick={() => {
                      insertWordArtAt(editor, p)
                      setDropdown(() => null)
                    }}
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
              disabled={!hasDoc}
              title={t('ribbonDropCap')}
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
              <div className="dropcap-menu">
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
                    title={opt.desc}
                    onClick={() => {
                      setDropdown(() => null)
                      const newDropCap = opt.val
                        ? JSON.stringify({ type: opt.val, lines: 3 })
                        : null
                      setParaAttrs(editor, { dropCap: newDropCap })
                    }}
                  >
                    <span className="dropcap-preview">{opt.val ? 'A' : '—'}</span>
                    <span>{opt.label}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              title={t('ribbonFieldTip')}
              onClick={() => toggleDropdown(setDropdown, 'field')}
            >
              <span className="rb-big-icon" style={{ fontSize: 15, lineHeight: 1.2 }}>
                {'{ }'}
                <IconCaret />
              </span>
              <span>{t('ribbonField')}</span>
            </button>
            {dropdown === 'field' && (
              <div className="layout-menu">
                {(
                  [
                    ['DATE', t('ribbonFieldDate')],
                    ['TIME', t('ribbonFieldTime')],
                    ['PAGE', t('ribbonFieldPage')],
                    ['NUMPAGES', t('ribbonFieldNumPages')],
                    ['FILENAME', t('ribbonFieldFileName')],
                  ] as const
                ).map(([instr, label]) => (
                  <button
                    key={instr}
                    onClick={() => {
                      onInsertField(instr)
                      setDropdown(() => null)
                    }}
                  >
                    {label}
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupText')}</div>
      </div>

      <div className="ribbon-sep" />

      {/* Word: Link/Bookmark/Cross-reference are small vertical buttons */}
      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-col">
            <button
              className="rb-small"
              disabled={!hasDoc}
              title={t('ribbonLinkTip')}
              onClick={() => setLinkOpen(true)}
            >
              <IconLink size={14} /> {t('ribbonLink')}
            </button>
            <button
              className="rb-small"
              disabled={!hasDoc}
              title={t('ribbonBookmarkTip')}
              onClick={() => setBookmarkOpen(true)}
            >
              <IconBook size={14} /> {t('ribbonBookmark')}
            </button>
            <button
              className="rb-small"
              disabled={!hasDoc}
              title={t('ribbonCrossRefTip')}
              onClick={() => setCrossRefOpen(true)}
            >
              <IconRefresh size={14} /> {t('ribbonCrossRef')}
            </button>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupLinks')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <button
            className="rb-big"
            disabled={!hasDoc}
            title={
              commentCount > 0
                ? t('ribbonViewCommentsTip', { count: commentCount })
                : t('ribbonViewCommentsNoneTip')
            }
            onClick={onShowComments}
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
              title={t('ribbonHeaderTip')}
              onClick={() => toggleDropdown(setDropdown, 'header')}
            >
              <span className="rb-big-icon">
                <IconHeader size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonHeader')}</span>
            </button>
            {dropdown === 'header' && (
              <HfEditor
                label={t('ribbonHeader')}
                current={header?.text ?? ''}
                onApply={(text) => onHeader({ text })}
                onClose={() => setDropdown(() => null)}
              />
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className={`rb-big ${footer?.text ? 'active' : ''}`}
              disabled={!hasDoc}
              title={t('ribbonFooterTip')}
              onClick={() => toggleDropdown(setDropdown, 'footer')}
            >
              <span className="rb-big-icon">
                <IconFooter size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonFooter')}</span>
            </button>
            {dropdown === 'footer' && (
              <HfEditor
                label={t('ribbonFooter')}
                current={footer?.text ?? ''}
                onApply={(text) => onFooter({ text, pageNumber: footer?.pageNumber ?? false })}
                onClose={() => setDropdown(() => null)}
              />
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className={`rb-big ${footer?.pageNumber ? 'active' : ''}`}
              disabled={!hasDoc}
              title={t('ribbonPageNumber')}
              onClick={() => toggleDropdown(setDropdown, 'pagenum')}
            >
              <span className="rb-big-icon">
                <IconPageNumber size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonPageNumber')}</span>
            </button>
            {dropdown === 'pagenum' && (
              <div className="layout-menu">
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
                  <button
                    key={label}
                    onClick={() => {
                      // Word gallery semantics: replace the area's content with a single page-number paragraph
                      const hf: HeaderFooter = {
                        text: '#',
                        pageNumber: true,
                        paras: [{ align, runs: [{ text: '#' }] }],
                      }
                      ;(kind === 'header' ? onHeader : onFooter)(hf)
                      setDropdown(() => null)
                    }}
                  >
                    {label}
                  </button>
                ))}
                <button
                  onClick={() => {
                    onPageNumFormat()
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonPnFormat')}
                </button>
                <button
                  onClick={() => {
                    if (footer?.pageNumber || footer?.text.includes('#')) {
                      onFooter({ text: (footer?.text ?? '').replace(/#/g, ''), pageNumber: false })
                    }
                    if (header?.pageNumber || header?.text.includes('#')) {
                      onHeader({ text: (header?.text ?? '').replace(/#/g, ''), pageNumber: false })
                    }
                    setDropdown(() => null)
                  }}
                >
                  {t('ribbonPnRemove')}
                </button>
              </div>
            )}
          </div>
          <div className="rb-col">
            <button
              className={`rb-small ${titlePg ? 'active' : ''}`}
              disabled={!hasDoc}
              title={t('ribbonDiffFirstPageTip')}
              onClick={() => onTitlePg(!titlePg)}
            >
              {titlePg ? <IconCheckboxChecked /> : <IconCheckbox />} {t('ribbonDiffFirstPage')}
            </button>
            <button
              className={`rb-small ${evenOddHf ? 'active' : ''}`}
              disabled={!hasDoc}
              title={t('ribbonDiffOddEvenTip')}
              onClick={() => onEvenOddHf(!evenOddHf)}
            >
              {evenOddHf ? <IconCheckboxChecked /> : <IconCheckbox />} {t('ribbonDiffOddEven')}
            </button>
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupHeaderFooter')}</div>
      </div>

      <div className="ribbon-sep" />

      <div className="ribbon-group">
        <div className="ribbon-group-items">
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              title={t('ribbonSymbolTip')}
              onClick={() => toggleDropdown(setDropdown, 'symbol')}
            >
              <span className="rb-big-icon">
                <IconSymbol size={BIG} />
                <IconCaret />
              </span>
              <span>{t('ribbonSymbol')}</span>
            </button>
            {dropdown === 'symbol' && (
              <div className="symbol-palette">
                {SYMBOLS.map((s) => (
                  <button
                    key={s}
                    className="symbol-cell"
                    onClick={() => {
                      editor.chain().focus().insertContent(s).run()
                      setDropdown(() => null)
                    }}
                  >
                    {s}
                  </button>
                ))}
              </div>
            )}
          </div>
          <div className="rb-split-wrap">
            <button
              className="rb-big"
              disabled={!hasDoc}
              title={t('ribbonEquationTip')}
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
                  setEquationOpen(true)
                }}
              />
            )}
          </div>
        </div>
        <div className="ribbon-group-label">{t('ribbonGroupSymbols')}</div>
      </div>

      {linkOpen && <LinkInsertModal editor={editor} onClose={() => setLinkOpen(false)} />}
      {equationOpen && <EquationModal editor={editor} onClose={() => setEquationOpen(false)} />}
      {bookmarkOpen && <BookmarkModal editor={editor} onClose={() => setBookmarkOpen(false)} />}
      {crossRefOpen && <CrossRefModal editor={editor} onClose={() => setCrossRefOpen(false)} />}
      {chartOpen && <ChartInsertModal editor={editor} onClose={() => setChartOpen(false)} />}
    </>
  )
}

/* ================= Design ================= */

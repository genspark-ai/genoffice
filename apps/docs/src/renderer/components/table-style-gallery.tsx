import { useEffect, useMemo, useState, type ReactNode } from 'react'
import {
  BUILT_IN_TABLE_STYLES,
  type BuiltInTableStyleGroup,
  type CellBorder,
  type StyleInfo,
  type TableLook,
  type TableStyleDisplay,
  type ThemeColors,
} from '@genoffice/docx-engine'
import { styledCell } from '../editor/table-style-apply'
import { loadBuiltInTableStyleInfos } from '../editor/built-in-table-style-infos'
import { useI18n, type StringKey } from '../i18n/locale'
import { IconCaret } from './icons'

// Previews draw document data: a style's own colors on white paper, as Word's
// gallery does in every UI theme.
const PAPER = '#FFFFFF'
const INK = '#000000'
const hex = (color: string | undefined, fallback: string) =>
  !color || color === 'auto' ? fallback : `#${color}`

/** a line's look in a preview: Word's gallery draws every style at a hairline or two */
function lineStroke(line: CellBorder): { width: number; dash?: string; double: boolean } {
  const width = (line.szEighths ?? 4) >= 12 ? 1.6 : 1
  const dash =
    line.style === 'dotted'
      ? '1 1.5'
      : line.style === 'dashed' || line.style === 'dashSmallGap'
        ? '3 2'
        : line.style === 'dotDash'
          ? '3 1.5 1 1.5'
          : line.style === 'dotDotDash'
            ? '3 1.5 1 1.5 1 1.5'
            : undefined
  return { width, dash, double: line.style === 'double' }
}

const ROWS = 5
const COLS = 5
const CELL_W = 13
const CELL_H = 9

/** a miniature 5×5 table in the style, under the table's current style options */
export function TableStyleThumb({
  display,
  look,
}: {
  display: TableStyleDisplay | undefined
  look: TableLook
}) {
  const cells: ReactNode[] = []
  const lines: ReactNode[] = []
  const textColor = (row: number, col: number) => {
    const cond =
      (look.firstRow && row === 0 && display?.firstRow) ||
      (look.lastRow && row === ROWS - 1 && display?.lastRow) ||
      (look.firstColumn && col === 0 && display?.firstCol) ||
      (look.lastColumn && col === COLS - 1 && display?.lastCol) ||
      undefined
    return {
      color: hex(cond?.color ?? display?.wholeTable?.color, INK),
      bold: !!(cond?.bold ?? display?.wholeTable?.bold),
    }
  }
  for (let row = 0; row < ROWS; row++) {
    for (let col = 0; col < COLS; col++) {
      const at = { row, rows: ROWS, col, span: 1, cols: COLS }
      const { fill, borders } = styledCell(display, look, at)
      const x = col * CELL_W
      const y = row * CELL_H
      const key = `${row}-${col}`
      cells.push(
        <rect
          key={`f${key}`}
          x={x}
          y={y}
          width={CELL_W}
          height={CELL_H}
          fill={hex(fill ?? undefined, PAPER)}
        />,
      )
      const text = textColor(row, col)
      cells.push(
        <line
          key={`t${key}`}
          x1={x + 3}
          x2={x + CELL_W - 3}
          y1={y + CELL_H / 2}
          y2={y + CELL_H / 2}
          stroke={text.color}
          strokeWidth={text.bold ? 1.6 : 0.9}
        />,
      )
      // each cell draws its top and left edge, the last row/column its far ones too
      const edges: Array<['top' | 'left' | 'bottom' | 'right', number, number, number, number]> = [
        ['top', x, y, x + CELL_W, y],
        ['left', x, y, x, y + CELL_H],
      ]
      if (row === ROWS - 1) edges.push(['bottom', x, y + CELL_H, x + CELL_W, y + CELL_H])
      if (col === COLS - 1) edges.push(['right', x + CELL_W, y, x + CELL_W, y + CELL_H])
      for (const [side, x1, y1, x2, y2] of edges) {
        const line = borders?.[side]
        if (!line || line.style === 'none' || line.style === 'nil') continue
        const stroke = lineStroke(line)
        const color = hex(line.color, INK)
        const offsets = stroke.double ? [-0.9, 0.9] : [0]
        for (const offset of offsets) {
          const horizontal = y1 === y2
          lines.push(
            <line
              key={`b${key}${side}${offset}`}
              x1={horizontal ? x1 : x1 + offset}
              x2={horizontal ? x2 : x2 + offset}
              y1={horizontal ? y1 + offset : y1}
              y2={horizontal ? y2 + offset : y2}
              stroke={color}
              strokeWidth={stroke.double ? 0.6 : stroke.width}
              strokeDasharray={stroke.dash}
            />,
          )
        }
      }
    }
  }
  return (
    <svg
      className="tbl-thumb-svg"
      width={COLS * CELL_W + 2}
      height={ROWS * CELL_H + 2}
      viewBox={`-1 -1 ${COLS * CELL_W + 2} ${ROWS * CELL_H + 2}`}
      aria-hidden="true"
    >
      {cells}
      {lines}
    </svg>
  )
}

interface GalleryEntry {
  styleId: string
  name: string
  group: BuiltInTableStyleGroup | 'custom'
  display: TableStyleDisplay | undefined
}

const GROUP_LABELS: Record<GalleryEntry['group'], StringKey> = {
  custom: 'ribbonTableStylesCustom',
  plain: 'ribbonTableStylesPlain',
  grid: 'ribbonTableStylesGrid',
  list: 'ribbonTableStylesList',
}

const PAGE = 6

/**
 * Table Design ▸ Table Styles: a strip of previews with Word's row arrows and
 * a ▾ that opens every style, grouped like Word's gallery (the document's own
 * styles first), with Clear at the bottom.
 */
export function TableStyleGallery({
  styles,
  themeColors,
  look,
  activeStyleId,
  open,
  onOpenChange,
  onPick,
}: {
  styles: Map<string, StyleInfo> | undefined
  themeColors: ThemeColors | null
  look: TableLook
  activeStyleId: string | null
  open: boolean
  onOpenChange: (open: boolean) => void
  /** a style id, or null to clear the table's style */
  onPick: (styleId: string | null) => void
}) {
  const { t } = useI18n()
  const [builtIns, setBuiltIns] = useState<Map<string, StyleInfo> | null>(null)
  useEffect(() => {
    let live = true
    void loadBuiltInTableStyleInfos(themeColors).then((infos) => {
      if (live) setBuiltIns(infos)
    })
    return () => {
      live = false
    }
  }, [themeColors])

  const entries = useMemo<GalleryEntry[]>(() => {
    const builtInIds = new Set(BUILT_IN_TABLE_STYLES.map((style) => style.styleId))
    const custom = [...(styles?.values() ?? [])]
      .filter(
        (info) =>
          info.type === 'table' && info.styleId !== 'TableNormal' && !builtInIds.has(info.styleId),
      )
      .map((info) => ({
        styleId: info.styleId,
        name: info.name,
        group: 'custom' as const,
        display: info.tableDisplay,
      }))
    const builtIn = BUILT_IN_TABLE_STYLES.map((style) => ({
      styleId: style.styleId,
      name: style.name,
      group: style.group,
      // a style the document already has is previewed as the document defines it
      display:
        styles?.get(style.styleId)?.tableDisplay ?? builtIns?.get(style.styleId)?.tableDisplay,
    }))
    return [...custom, ...builtIn]
  }, [styles, builtIns])

  const activeIndex = entries.findIndex((entry) => entry.styleId === activeStyleId)
  const [page, setPage] = useState(0)
  useEffect(() => {
    if (activeIndex >= 0) setPage(Math.floor(activeIndex / PAGE))
  }, [activeIndex])
  const pages = Math.max(1, Math.ceil(entries.length / PAGE))
  const shown = entries.slice(page * PAGE, page * PAGE + PAGE)

  const thumb = (entry: GalleryEntry) => (
    <button
      key={entry.styleId}
      className={`tbl-thumb${entry.styleId === activeStyleId ? ' active' : ''}`}
      data-tip={entry.name}
      aria-label={entry.name}
      aria-pressed={entry.styleId === activeStyleId}
      onClick={() => {
        onPick(entry.styleId)
        onOpenChange(false)
      }}
    >
      <TableStyleThumb display={entry.display} look={look} />
    </button>
  )

  return (
    <div className="tbl-gallery">
      <div className="tbl-gallery-strip">{shown.map(thumb)}</div>
      <div className="tbl-gallery-nav">
        <button
          className="tbl-gallery-step"
          aria-label={t('ribbonTableStylesPrev')}
          data-tip={t('ribbonTableStylesPrev')}
          disabled={page === 0}
          onClick={() => setPage((p) => Math.max(0, p - 1))}
        >
          <span className="tbl-gallery-up">
            <IconCaret size={10} />
          </span>
        </button>
        <button
          className="tbl-gallery-step"
          aria-label={t('ribbonTableStylesNext')}
          data-tip={t('ribbonTableStylesNext')}
          disabled={page >= pages - 1}
          onClick={() => setPage((p) => Math.min(pages - 1, p + 1))}
        >
          <IconCaret size={10} />
        </button>
        <button
          className={`tbl-gallery-step tbl-gallery-more${open ? ' active' : ''}`}
          aria-label={t('ribbonTableStylesMore')}
          data-tip={t('ribbonTableStylesMore')}
          aria-expanded={open}
          onClick={() => onOpenChange(!open)}
        >
          <IconCaret size={10} />
        </button>
      </div>
      {open && (
        <div data-rb-panel="" className="tbl-gallery-panel">
          {(['custom', 'plain', 'grid', 'list'] as const).map((group) => {
            const inGroup = entries.filter((entry) => entry.group === group)
            if (inGroup.length === 0) return null
            return (
              <section key={group}>
                <div className="tbl-gallery-heading">{t(GROUP_LABELS[group])}</div>
                <div className="tbl-gallery-grid">{inGroup.map(thumb)}</div>
              </section>
            )
          })}
          <div className="tbl-gallery-sep" />
          <button
            className="tbl-gallery-clear"
            onClick={() => {
              onPick(null)
              onOpenChange(false)
            }}
          >
            {t('ribbonTableStyleClear')}
          </button>
        </div>
      )}
    </div>
  )
}

/** a line sample for the pen pickers */
export function LineSample({ pen, width = 64 }: { pen: CellBorder; width?: number }) {
  const px = Math.max(1, Math.min(6, ((pen.szEighths ?? 4) / 8) * 1.333))
  const dash =
    pen.style === 'dotted'
      ? `${px} ${px * 1.5}`
      : pen.style === 'dashed'
        ? `${px * 4} ${px * 2}`
        : pen.style === 'dotDash'
          ? `${px * 4} ${px * 1.5} ${px} ${px * 1.5}`
          : pen.style === 'dotDotDash'
            ? `${px * 4} ${px * 1.5} ${px} ${px * 1.5} ${px} ${px * 1.5}`
            : undefined
  const color = 'currentColor'
  const y = 6
  return (
    <svg className="tbl-line-sample" width={width} height={12} aria-hidden="true">
      {pen.style === 'double' ? (
        <>
          <line x1={0} x2={width} y1={y - 2} y2={y - 2} stroke={color} strokeWidth={1} />
          <line x1={0} x2={width} y1={y + 2} y2={y + 2} stroke={color} strokeWidth={1} />
        </>
      ) : pen.style === 'triple' ? (
        <>
          <line x1={0} x2={width} y1={y - 3} y2={y - 3} stroke={color} strokeWidth={1} />
          <line x1={0} x2={width} y1={y} y2={y} stroke={color} strokeWidth={1} />
          <line x1={0} x2={width} y1={y + 3} y2={y + 3} stroke={color} strokeWidth={1} />
        </>
      ) : (
        <line
          x1={0}
          x2={width}
          y1={y}
          y2={y}
          stroke={color}
          strokeWidth={pen.style === 'thick' ? Math.max(px, 2) : px}
          strokeDasharray={dash}
        />
      )}
    </svg>
  )
}

/** Word's pen weights, in eighths of a point, with the labels its picker shows */
export const PEN_WEIGHTS: ReadonlyArray<[number, string]> = [
  [2, '¼'],
  [4, '½'],
  [6, '¾'],
  [8, '1'],
  [12, '1½'],
  [18, '2¼'],
  [24, '3'],
  [36, '4½'],
  [48, '6'],
]

export const PEN_STYLES = [
  'single',
  'dotted',
  'dashed',
  'dotDash',
  'dotDotDash',
  'double',
  'triple',
  'thick',
] as const

/** Office theme accents for Border Styles when the document has no theme */
const OFFICE_ACCENTS = ['156082', 'E97132', '196B24', '0F9ED5', 'A02B93', '4EA72E']

/** Border Styles ▾: Word's theme borders, three pens in black and each accent */
export function themeBorderPens(themeColors: ThemeColors | null): CellBorder[][] {
  const accents = [1, 2, 3, 4, 5, 6].map(
    (n, i) => themeColors?.[`accent${n}` as keyof ThemeColors] ?? OFFICE_ACCENTS[i],
  )
  const colors = ['000000', ...accents]
  return [
    colors.map((color) => ({ style: 'single', szEighths: 4, color })),
    colors.map((color) => ({ style: 'single', szEighths: 12, color })),
    colors.map((color) => ({ style: 'double', szEighths: 4, color })),
  ]
}

/**
 * Ribbon: tab bar + grouped buttons. Same mechanism as the apps/docs Ribbon
 * (local state switches tabs, .ribbon-body dispatches); content is trimmed to slide capabilities,
 * unimplemented items are grayed placeholders.
 */
import React, {
  useCallback,
  useEffect,
  useLayoutEffect,
  useRef,
  useState,
  type ReactNode,
} from 'react'
import type { AnimEffectKind, AnimTrigger, TransitionKind } from '../../shared/ipc'
import type { ChartStyleInfo } from '@genoffice/pptx-render'
import { useDismissablePopover } from '@genoffice/ui'
import { ICON_COLORS } from '../insert-presets'
import { THEME_PRESETS, type SlideThemePreset } from '../themes'
import { restoreEditSelection } from '../TextEditOverlay'
import { armColorInput, toPickerHex } from '../color-input'
import { TABLE_SHADING_COLORS } from './table-shading-colors'
import { useI18n, type StringKey } from '../i18n/locale'
import {
  IconSlideMaster,
  IconBullets,
  IconArrangeAll,
  IconComment,
  IconCursor,
  IconCustomShow,
  IconEraser,
  IconHideSlide,
  IconNavPane,
  IconOutlineView,
  IconPageColor,
  IconApplyAll,
  IconPageSize,
  IconPlayCurrent,
  IconPlayFromStart,
  IconPresenterView,
  IconPrintLayout,
  IconReadMode,
  IconRecord,
  IconRemoveBg,
  IconRedo,
  IconRehearse,
  IconSave,
  IconSetupShow,
  IconSparkle,
  IconUndo,
  IconFitWindow,
  IconZoom100,
  IconZoomIn,
  IconZoomOut,
  IconSwitchRowCol,
  IconEditChartData,
  IconChangeChartType,
  AnimEffectIcon,
  IconTransNone,
  IconTransMorph,
  IconTransFade,
  IconTransPush,
  IconTransWipe,
  IconTransSplit,
  IconTransCircle,
  IconTransCover,
  IconTransPull,
  IconTransDissolve,
  IconTransZoom,
  IconTransRandom,
  IconAnimStar,
  IconAnimNone,
  IconPageBorders,
  IconNoneX,
  IconPathRight,
  IconPathDown,
  IconPathDiagonal,
  IconPathCircle,
  IconPathZigzag,
  IconShapeStyle,
  IconFillColor,
} from './icons'
// brand-supplied Review AI icon art (44px = 22px @2x), color baked in
import iconSpelling from '../assets/icon-spelling.png'
import iconTranslate from '../assets/icon-translate.png'
import iconTransparency from '../assets/icon-transparency.png'
import iconCrop from '../assets/icon-crop.png'
import { ChartTypeDialog } from './ChartTypeDialog'
import {
  BIG,
  Group,
  RbCaret,
  RIBBON_SHAPE_STYLES,
  TEXT_COLORS,
  closeSiblingPanels,
  type Props,
  type RibbonPanelKey,
  type RibbonTabCtx,
} from './ribbon-shared'
export type { FormatCmd, SlidesViewMode } from './ribbon-shared'
import type { FormatCmd } from './ribbon-shared'
import { RibbonHomeTab } from './RibbonHomeTab'
import { RibbonInsertTab } from './RibbonInsertTab'
import { autoContextTabForElement, contextTabForElement, type ContextTab } from './context-tabs'

const IS_MAC = navigator.platform.toLowerCase().includes('mac')
/** shell tab mode: the tab strip above owns traffic lights / caption buttons */
const IN_TAB = new URLSearchParams(window.location.search).get('mode') === 'tab'

type MainTab =
  | 'file'
  | 'home'
  | 'insert'
  | 'draw'
  | 'design'
  | 'transitions'
  | 'animations'
  | 'slideShow'
  | 'review'
  | 'view'

// Mac has no "File" tab (file operations go through the native menu), Windows does
const TABS: readonly MainTab[] = IS_MAC
  ? ['home', 'insert', 'draw', 'design', 'transitions', 'animations', 'slideShow', 'review', 'view']
  : [
      'file',
      'home',
      'insert',
      'draw',
      'design',
      'transitions',
      'animations',
      'slideShow',
      'review',
      'view',
    ]

const TAB_LABEL: Record<MainTab | ContextTab, StringKey> = {
  file: 'ribbonTabFile',
  home: 'ribbonTabHome',
  insert: 'ribbonTabInsert',
  draw: 'ribbonTabDraw',
  design: 'ribbonTabDesign',
  transitions: 'ribbonTabTransitions',
  animations: 'ribbonTabAnimations',
  slideShow: 'ribbonTabSlideShow',
  review: 'ribbonTabReview',
  view: 'ribbonTabView',
  tableDesign: 'ribbonTabTableDesign',
  chartDesign: 'ribbonTabChartDesign',
  pictureFormat: 'ribbonTabPictureFormat',
  shapeFormat: 'ribbonTabShapeFormat',
}

// display names only — tp.name stays as written into theme*.xml
const THEME_NAME: Record<string, StringKey> = {
  office: 'ribbonThemeOffice',
  ember: 'ribbonThemeEmber',
  indigo: 'ribbonThemeIndigo',
  forest: 'ribbonThemeForest',
  cream: 'ribbonThemeCream',
  rose: 'ribbonThemeRose',
  graphite: 'ribbonThemeGraphite',
  midnight: 'ribbonThemeMidnight',
}
const themeDisplayName = (tp: SlideThemePreset, t: (key: StringKey) => string): string =>
  THEME_NAME[tp.id] ? t(THEME_NAME[tp.id]) : tp.name

/** Translation target languages (for AI proofread/translate presets) */
const TRANSLATE_TARGETS: StringKey[] = [
  'ribbonLangEnglish',
  'ribbonLangSimplifiedChinese',
  'ribbonLangTraditionalChinese',
  'ribbonLangJapanese',
  'ribbonLangKorean',
  'ribbonLangFrench',
  'ribbonLangGerman',
  'ribbonLangSpanish',
]

/** One-time "AI rewrites the whole document" acknowledgement */
const AI_REWRITE_ACK_KEY = 'slides-ai-rewrite-ack'

// Draw tab palettes/pen widths (same as apps/docs DrawTab)
const INK_COLORS = [
  '000000',
  'C00000',
  'FF0000',
  'FFC000',
  'FFFF00',
  '92D050',
  '00B050',
  '00B0F0',
  '0070C0',
  '7030A0',
]
const PEN_WIDTHS = [1, 2, 3.5, 5]
const HIGHLIGHTER_WIDTHS = [6, 10, 16]

/** Draw-tab pen gallery presets (tray of ready pens) */
interface PenPreset {
  kind: 'pen' | 'highlighter'
  color: string
  width: number
}

const DEFAULT_PEN_PRESETS: PenPreset[] = [
  { kind: 'pen', color: '000000', width: 2 },
  { kind: 'pen', color: 'FF0000', width: 2 },
  { kind: 'pen', color: '0070C0', width: 2 },
  { kind: 'highlighter', color: 'FFFF00', width: 10 },
  { kind: 'highlighter', color: '00B050', width: 10 },
]

/** Pen thumbnail hanging tip-down in the tray */
function PenThumb({ kind, color }: { kind: PenPreset['kind']; color: string }) {
  const c = `#${color}`
  return kind === 'pen' ? (
    <svg width="24" height="56" viewBox="0 0 24 56" aria-hidden="true">
      <rect x="5" y="0" width="14" height="34" rx="2.5" fill={c} />
      <path d="M5 34h14l-3.5 8h-7Z" fill={c} opacity="0.85" />
      <rect x="8.6" y="41" width="6.8" height="3.4" rx="1.2" fill="#fff" opacity="0.9" />
      <path d="M8.5 44.5h7L12 55Z" fill={c} />
    </svg>
  ) : (
    <svg width="26" height="56" viewBox="0 0 26 56" aria-hidden="true">
      <rect x="3" y="0" width="20" height="36" rx="2.5" fill={c} />
      <path d="M6 36h14l-2.5 9h-8Z" fill={c} opacity="0.8" />
      <path d="M9 45h8l-1.5 8h-5Z" fill={c} />
    </svg>
  )
}

const TRANSITIONS: Array<{ kind: TransitionKind; label: StringKey; icon: React.ReactElement }> = [
  { kind: 'none', label: 'ribbonNone', icon: <IconTransNone size={BIG} /> },
  { kind: 'morph', label: 'ribbonTransMorph', icon: <IconTransMorph size={BIG} /> },
  { kind: 'fade', label: 'ribbonTransFade', icon: <IconTransFade size={BIG} /> },
  { kind: 'push', label: 'ribbonTransPush', icon: <IconTransPush size={BIG} /> },
  { kind: 'wipe', label: 'ribbonTransWipe', icon: <IconTransWipe size={BIG} /> },
  { kind: 'split', label: 'ribbonTransSplit', icon: <IconTransSplit size={BIG} /> },
  { kind: 'circle', label: 'ribbonTransCircle', icon: <IconTransCircle size={BIG} /> },
  { kind: 'cover', label: 'ribbonTransCover', icon: <IconTransCover size={BIG} /> },
  { kind: 'pull', label: 'ribbonTransPull', icon: <IconTransPull size={BIG} /> },
  { kind: 'dissolve', label: 'ribbonTransDissolve', icon: <IconTransDissolve size={BIG} /> },
  { kind: 'zoom', label: 'ribbonTransZoom', icon: <IconTransZoom size={BIG} /> },
  { kind: 'random', label: 'ribbonTransRandom', icon: <IconTransRandom size={BIG} /> },
]

/** Animation effect gallery (entrance/emphasis/exit; icons drawn per effect, colored by class). */
const ANIM_EFFECTS: Array<{
  kind: AnimEffectKind
  label: StringKey
  cls: 'entr' | 'emph' | 'exit'
}> = [
  { kind: 'appear', label: 'ribbonAnimAppear', cls: 'entr' },
  { kind: 'fade', label: 'ribbonAnimFade', cls: 'entr' },
  { kind: 'flyIn', label: 'ribbonAnimFlyIn', cls: 'entr' },
  { kind: 'wipe', label: 'ribbonAnimWipe', cls: 'entr' },
  { kind: 'wipeDown', label: 'ribbonAnimWipeDown', cls: 'entr' },
  { kind: 'splitIn', label: 'ribbonAnimSplit', cls: 'entr' },
  { kind: 'bounce', label: 'ribbonAnimBounce', cls: 'entr' },
  { kind: 'flipIn', label: 'ribbonAnimFlip', cls: 'entr' },
  { kind: 'zoom', label: 'ribbonAnimZoom', cls: 'entr' },
  { kind: 'pulse', label: 'ribbonAnimPulse', cls: 'emph' },
  { kind: 'spin', label: 'ribbonAnimSpin', cls: 'emph' },
  { kind: 'grow', label: 'ribbonAnimGrowShrink', cls: 'emph' },
  { kind: 'teeter', label: 'ribbonAnimTeeter', cls: 'emph' },
  { kind: 'disappear', label: 'ribbonAnimDisappear', cls: 'exit' },
  { kind: 'fadeOut', label: 'ribbonAnimFadeOut', cls: 'exit' },
  { kind: 'flyOut', label: 'ribbonAnimFlyOut', cls: 'exit' },
  { kind: 'wipeOut', label: 'ribbonAnimWipeOut', cls: 'exit' },
  { kind: 'shrink', label: 'ribbonAnimShrinkTurn', cls: 'exit' },
  { kind: 'zoomOut', label: 'ribbonAnimZoomOut', cls: 'exit' },
]

const ANIM_CLS_TITLE: Record<'entr' | 'emph' | 'exit', StringKey> = {
  entr: 'ribbonAnimEntrance',
  emph: 'ribbonAnimEmphasis',
  exit: 'ribbonAnimExit',
}

/** Motion path presets (path coordinates 0..1 relative to slide size, matching OOXML animMotion). */
const MOTION_PATHS: Array<{ label: StringKey; icon: React.ReactElement; path: string }> = [
  { label: 'ribbonPathLineRight', icon: <IconPathRight size={BIG} />, path: 'M 0 0 L 0.25 0' },
  { label: 'ribbonPathLineDown', icon: <IconPathDown size={BIG} />, path: 'M 0 0 L 0 0.25' },
  { label: 'ribbonPathDiagonal', icon: <IconPathDiagonal size={BIG} />, path: 'M 0 0 L 0.25 0.25' },
  {
    label: 'ribbonPathCircle',
    icon: <IconPathCircle size={BIG} />,
    path:
      'M 0 0 C 0.069 0 0.125 0.056 0.125 0.125 C 0.125 0.194 0.069 0.25 0 0.25 ' +
      'C -0.069 0.25 -0.125 0.194 -0.125 0.125 C -0.125 0.056 -0.069 0 0 0 Z',
  },
  {
    label: 'ribbonPathZigzag',
    icon: <IconPathZigzag size={BIG} />,
    path: 'M 0 0 L 0.12 -0.12 L 0.24 0.12 L 0.36 0',
  },
]

// ── Constants + helper components used by contextual tabs ─────────────────────

/** Mini table thumbnail of a preset style (colors match the engine-side TABLE_STYLE_PRESETS literals) */
function TableMiniPreview({
  header,
  band,
  rowLine,
  gridLine,
  outline,
}: {
  /** Header row background color */
  header?: string
  /** Banded stripe color (rows 2/4) */
  band?: string
  /** Horizontal line color between rows */
  rowLine?: string
  /** All-gridlines color (horizontal + vertical) */
  gridLine?: string
  /** Outer border color */
  outline?: string
}) {
  const W = 44
  const H = 26
  const rows = 4
  const cols = 3
  const rh = H / rows
  const cw = W / cols
  const hLines = [1, 2, 3].map((i) => i * rh)
  const line = gridLine ?? rowLine
  return (
    <svg width={W} height={H} shapeRendering="crispEdges" aria-hidden>
      <rect x={0} y={0} width={W} height={H} fill="#fff" />
      {band &&
        [1, 3].map((r) => <rect key={r} x={0} y={r * rh} width={W} height={rh} fill={band} />)}
      {header && <rect x={0} y={0} width={W} height={rh} fill={header} />}
      {line &&
        hLines.map((y) => (
          <line key={y} x1={0} y1={y} x2={W} y2={y} stroke={line} strokeWidth={1} />
        ))}
      {gridLine &&
        [1, 2].map((c) => (
          <line key={c} x1={c * cw} y1={0} x2={c * cw} y2={H} stroke={gridLine} strokeWidth={1} />
        ))}
      <rect
        x={0.5}
        y={0.5}
        width={W - 1}
        height={H - 1}
        fill="none"
        stroke={outline ?? '#E3E3E3'}
        strokeWidth={1}
      />
    </svg>
  )
}

/** UI descriptions of the 8 table preset styles */
const TABLE_STYLE_PRESETS_UI: Array<{
  key: string
  label: StringKey
  preview: { header?: string; band?: string; rowLine?: string; gridLine?: string; outline?: string }
}> = [
  { key: 'none', label: 'ribbonTableStyleNone', preview: {} },
  {
    key: 'lightGrid',
    label: 'ribbonTableStyleLightGrid',
    preview: { gridLine: '#BFBFBF', outline: '#BFBFBF' },
  },
  {
    key: 'zebraBlue',
    label: 'ribbonTableStyleZebraBlue',
    preview: { header: '#4472C4', band: '#D6E4F0', outline: '#C9D8EA' },
  },
  {
    key: 'zebraGray',
    label: 'ribbonTableStyleZebraGray',
    preview: { header: '#595959', band: '#EDEDED', outline: '#D0D0D0' },
  },
  {
    key: 'headerDarkBlue',
    label: 'ribbonTableStyleHeaderDarkBlue',
    preview: { header: '#1F3864', rowLine: '#D9D9D9', outline: '#D0D0D0' },
  },
  {
    key: 'headerOrange',
    label: 'ribbonTableStyleHeaderOrange',
    preview: { header: '#ED7D31', rowLine: '#D9D9D9', outline: '#D0D0D0' },
  },
  { key: 'noBorder', label: 'ribbonTableStyleNoBorder', preview: { band: '#F2F2F2' } },
  {
    key: 'fullBorder',
    label: 'ribbonTableStyleFullBorder',
    preview: { gridLine: '#595959', outline: '#595959' },
  },
]

/** Chart style presets (legend/gridlines/data labels/bar width combinations, applied to the same-named EditChartOp fields) */
interface ChartStylePreset {
  key: string
  label: StringKey
  style: {
    legendPos: 'b' | 't' | 'r' | 'l' | 'none'
    dataLabels: boolean
    gridlines: boolean
    gapWidthPct: number
  }
}
const CHART_STYLE_PRESETS: ChartStylePreset[] = [
  {
    key: 'classic',
    label: 'ribbonChartStyleClassic',
    style: { legendPos: 'b', dataLabels: false, gridlines: false, gapWidthPct: 150 },
  },
  {
    key: 'grid',
    label: 'ribbonChartStyleGrid',
    style: { legendPos: 'b', dataLabels: false, gridlines: true, gapWidthPct: 150 },
  },
  {
    key: 'labeled',
    label: 'ribbonDataLabels',
    style: { legendPos: 'b', dataLabels: true, gridlines: false, gapWidthPct: 150 },
  },
  {
    key: 'detail',
    label: 'ribbonChartStyleDetail',
    style: { legendPos: 'b', dataLabels: true, gridlines: true, gapWidthPct: 150 },
  },
  {
    key: 'minimal',
    label: 'ribbonChartStyleMinimal',
    style: { legendPos: 'none', dataLabels: false, gridlines: false, gapWidthPct: 150 },
  },
  {
    key: 'minimal-labeled',
    label: 'ribbonChartStyleMinimalLabeled',
    style: { legendPos: 'none', dataLabels: true, gridlines: false, gapWidthPct: 150 },
  },
  {
    key: 'bold',
    label: 'ribbonChartStyleBold',
    style: { legendPos: 'b', dataLabels: false, gridlines: true, gapWidthPct: 50 },
  },
  {
    key: 'slim',
    label: 'ribbonChartStyleSlim',
    style: { legendPos: 'b', dataLabels: false, gridlines: true, gapWidthPct: 300 },
  },
  {
    key: 'legend-top',
    label: 'ribbonChartStyleLegendTop',
    style: { legendPos: 't', dataLabels: false, gridlines: true, gapWidthPct: 150 },
  },
  {
    key: 'legend-right',
    label: 'ribbonChartStyleLegendRight',
    style: { legendPos: 'r', dataLabels: false, gridlines: true, gapWidthPct: 150 },
  },
]

/** Whether the current chart style matches a preset (bar width compared only for bar-family charts). */
function chartPresetActive(info: ChartStyleInfo | null | undefined, p: ChartStylePreset): boolean {
  if (!info) return false
  const s = p.style
  const barKind =
    info.kind === 'bar' ||
    info.kind === 'bar3D' ||
    info.kind === 'barStacked' ||
    info.kind === 'comboBarLine'
  return (
    info.legendPos === s.legendPos &&
    info.dataLabels === s.dataLabels &&
    info.gridlines === s.gridlines &&
    (!barKind || (info.gapWidthPct ?? 150) === s.gapWidthPct)
  )
}

/** Style preset thumbnail (draws a bar/line/pie sample per the current chart type). */
function ChartStyleThumb({
  kind,
  style,
}: {
  kind?: ChartStyleInfo['kind']
  style: ChartStylePreset['style']
}) {
  const W = 64
  const H = 40
  const C1 = '#4472C4'
  const C2 = '#ED7D31'
  const plot = { x: 5, y: 4, w: W - 10, h: H - 9 }
  if (style.legendPos === 'b') plot.h -= 6
  else if (style.legendPos === 't') {
    plot.y += 7
    plot.h -= 7
  } else if (style.legendPos === 'r') plot.w -= 13
  else if (style.legendPos === 'l') {
    plot.x += 13
    plot.w -= 13
  }
  const els: ReactNode[] = []
  const family =
    kind === 'line' || kind === 'area' || kind === 'scatter' || kind === 'radar'
      ? 'line'
      : kind === 'pie' || kind === 'pie3D' || kind === 'doughnut'
        ? 'pie'
        : 'bar'
  if (style.gridlines && family !== 'pie') {
    for (let i = 1; i <= 3; i++) {
      const y = plot.y + (plot.h * i) / 4
      els.push(
        <line
          key={`g${i}`}
          x1={plot.x}
          y1={y}
          x2={plot.x + plot.w}
          y2={y}
          stroke="#DADADA"
          strokeWidth={0.8}
        />,
      )
    }
  }
  if (family === 'bar') {
    const vals: Array<[number, number]> = [
      [0.55, 0.35],
      [0.8, 0.5],
      [0.65, 0.9],
    ]
    const slot = plot.w / 3
    const barW = Math.min(slot / (2 + style.gapWidthPct / 100), slot * 0.45)
    vals.forEach(([a, b], i) => {
      const gx = plot.x + i * slot + (slot - barW * 2) / 2
      const draw = (v: number, off: number, color: string, key: string) => {
        const h = plot.h * v
        els.push(
          <rect
            key={key}
            x={gx + off}
            y={plot.y + plot.h - h}
            width={barW}
            height={h}
            fill={color}
          />,
        )
        if (style.dataLabels)
          els.push(
            <circle
              key={`${key}d`}
              cx={gx + off + barW / 2}
              cy={plot.y + plot.h - h - 2.5}
              r={1.1}
              fill="#777"
            />,
          )
      }
      draw(a, 0, C1, `b${i}a`)
      draw(b, barW, C2, `b${i}b`)
    })
  } else if (family === 'line') {
    const mk = (vals: number[], color: string, key: string) => {
      const pts = vals.map(
        (v, i) => [plot.x + (plot.w * (i + 0.5)) / vals.length, plot.y + plot.h * (1 - v)] as const,
      )
      els.push(
        <polyline
          key={key}
          points={pts.map((p) => p.join(',')).join(' ')}
          fill="none"
          stroke={color}
          strokeWidth={1.5}
        />,
      )
      if (style.dataLabels)
        pts.forEach((p, i) =>
          els.push(<circle key={`${key}d${i}`} cx={p[0]} cy={p[1] - 2.5} r={1.1} fill="#777" />),
        )
    }
    mk([0.3, 0.55, 0.45, 0.8], C1, 'l1')
    mk([0.15, 0.35, 0.6, 0.5], C2, 'l2')
  } else {
    const cx = plot.x + plot.w / 2
    const cy = plot.y + plot.h / 2
    const r = Math.min(plot.w, plot.h) / 2 - 1
    // Main circle + 120° sector (clockwise from 12 o'clock)
    els.push(<circle key="p1" cx={cx} cy={cy} r={r} fill={C1} />)
    const a = ((-90 + 120) * Math.PI) / 180
    els.push(
      <path
        key="p2"
        d={`M ${cx} ${cy} L ${cx} ${cy - r} A ${r} ${r} 0 0 1 ${cx + Math.cos(a) * r} ${cy + Math.sin(a) * r} Z`}
        fill={C2}
      />,
    )
    if (kind === 'doughnut')
      els.push(<circle key="ph" cx={cx} cy={cy} r={r * 0.45} fill="var(--surface, #fff)" />)
    if (style.dataLabels) {
      els.push(<circle key="pd1" cx={cx - r * 0.45} cy={cy + r * 0.2} r={1.1} fill="#fff" />)
      els.push(<circle key="pd2" cx={cx + r * 0.5} cy={cy - r * 0.3} r={1.1} fill="#fff" />)
    }
  }
  // Legend dots
  if (style.legendPos !== 'none') {
    const horiz = style.legendPos === 'b' || style.legendPos === 't'
    const ly = style.legendPos === 'b' ? H - 4 : style.legendPos === 't' ? 4 : H / 2 - 3
    const lx = style.legendPos === 'r' ? W - 9 : style.legendPos === 'l' ? 4 : W / 2 - 7
    els.push(<rect key="lg1" x={lx} y={ly - 1.5} width={5} height={3} fill={C1} />)
    els.push(
      <rect
        key="lg2"
        x={horiz ? lx + 9 : lx}
        y={horiz ? ly - 1.5 : ly + 4.5}
        width={5}
        height={3}
        fill={C2}
      />,
    )
  }
  return (
    <svg width={W} height={H} viewBox={`0 0 ${W} ${H}`} aria-hidden>
      {els}
    </svg>
  )
}

/** Chart color schemes */
const CHART_COLOR_SCHEME_UI: Array<{ key: string; label: StringKey; colors: string[] }> = [
  { key: 'default', label: 'ribbonSchemeDefault', colors: [] },
  { key: 'blue', label: 'ribbonSchemeBlue', colors: ['#2E75B6', '#4472C4', '#5B9BD5', '#70AD47'] },
  { key: 'warm', label: 'ribbonSchemeWarm', colors: ['#ED7D31', '#FFC000', '#FF0000', '#C55A11'] },
  { key: 'cool', label: 'ribbonSchemeCool', colors: ['#0070C0', '#00B0F0', '#00B0A0', '#7030A0'] },
  { key: 'mono', label: 'ribbonSchemeMono', colors: ['#404040', '#666666', '#888888', '#AAAAAA'] },
]

/** Table option toggle button (state can't be displayed, click only toggles). */
function TableToggleBtn({
  label,
  on,
  disabled,
  onClick,
  offClick,
}: {
  label: string
  on: boolean
  disabled?: boolean
  onClick: () => void
  offClick: () => void
}) {
  const { t } = useI18n()
  return (
    <button
      className={`rb-icon ${on ? 'active' : ''}`}
      disabled={disabled}
      data-tip={t(on ? 'ribbonToggleOffTip' : 'ribbonToggleOnTip', { name: label })}
      onClick={() => (on ? offClick() : onClick())}
    >
      {label}
    </button>
  )
}

function DisabledBig({ icon, label }: { icon: ReactNode; label: string }) {
  const { t } = useI18n()
  return (
    <button className="rb-big" disabled data-tip={t('ribbonNotSupported', { name: label })}>
      <span className="rb-big-icon">{icon}</span>
      <span>{label}</span>
    </button>
  )
}

/** per-tab priority for responsive collapse: when the ribbon
 * body overflows, these groups (in order) fold into a single dropdown button */
const COLLAPSE_ORDER: Record<string, string[]> = {
  animations: ['motionPaths', 'animation'],
}

/** Checkbox row for toggle commands (View tab's Show group) */
function RbCheck({
  label,
  on,
  disabled,
  title,
  onClick,
}: {
  label: string
  on: boolean
  disabled?: boolean
  title?: string
  onClick: () => void
}) {
  return (
    <button
      className={`rb-check${on ? ' on' : ''}`}
      disabled={disabled}
      data-tip={title}
      onClick={onClick}
    >
      <span className="rb-check-box">
        {on && (
          <svg width="12" height="12" viewBox="0 0 24 24" aria-hidden="true">
            <path
              d="M5 12.4 10 17.4l9-10.8"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.5"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </svg>
        )}
      </span>
      <span>{label}</span>
    </button>
  )
}

export function Ribbon({
  hasDoc,
  deckEmpty,
  canUndo,
  canRedo,
  dirty,
  editing,
  autoSave,
  onAutoSaveChange,
  onOpen,
  onSave,
  onUndo,
  onRedo,
  onSaveAs,
  onExportPdf,
  onPrint,
  onExportImages,
  onFormat,
  zoom,
  onZoom,
  showThumbs,
  onToggleThumbs,
  aiOpen,
  onToggleAi,
  onAiPreset,
  onInsert,
  onPickShape,
  onInsertImage,
  onFormatBackground,
  onApplyTheme,
  onAddSlide,
  onAddSlideWithLayout,
  onAddSection,
  layouts,
  layoutSize,
  formatOpen,
  onToggleFormat,
  hasSelection,
  hasTextSelection,
  canPaste,
  onCopy,
  onCut,
  onPaste,
  hasBrushFormat,
  brushMode,
  onFormatBrushClick,
  onFormatBrushDoubleClick,
  onTextColor,
  curBulletChar,
  curAlign,
  curFontFamily,
  curFontSizePt,
  curFontSizeMixed,
  onFontFamily,
  onFontSize,
  onAlign,
  onStrike,
  onTextToggle,
  onElementTextColor,
  onFindReplace,
  animByParagraph,
  onToggleAnimByParagraph,
  onSetLayout,
  onResetLayout,
  onSlideSize,
  slideSizeKey,
  onParagraphFormat,
  onInsertTable,
  transition,
  onTransition,
  selectedAnimEffect,
  timingAnim,
  onApplyAnimation,
  onAnimHoverPreview,
  onAnimHoverEnd,
  onAddAnimation,
  onApplyMotionPath,
  onAnimTiming,
  animPaneOpen,
  onToggleAnimPane,
  animCount,
  onAnimPreview,
  onSlideShow,
  onPresenterView,
  onCustomShow,
  onRehearse,
  currentHidden,
  onToggleHidden,
  inkTool,
  onInkTool,
  inkPen,
  onInkPen,
  inkHighlighter,
  onInkHighlighter,
  inkCount,
  onInkClearAll,
  viewMode,
  onViewMode,
  onSlideMaster,
  onZoomFit,
  showRuler,
  onToggleRuler,
  showGrid,
  onToggleGrid,
  showGuides,
  onToggleGuides,
  showNotes,
  onToggleNotes,
  commentsOpen,
  onToggleComments,
  onNewComment,
  commentCount,
  onInsertIcon,
  onInsertChart,
  onInsertSmartArt,
  onInsertWordArt,
  onInsertField,
  onOpenLink,
  onInsertZoom,
  slideCount,
  currentSlide,
  onOpenHeaderFooter,
  onOpenEquation,
  onInsertMedia,
  onInsertModel3d,
  recording,
  onToggleScreenRecord,
  contextElementType,
  contextElementId: _contextElementId,
  contextSlideIndex: _contextSlideIndex,
  contextChartStyle,
  chartColorSchemes,
  contextPictureCanCutout,
  contextPictureStroke,
  onPictureStroke,
  onShapeStyle,
  onShapeFill,
  onPictureCrop,
  cropActive,
  onPictureOpacity,
  onPictureCutout,
  onEditTableStyle,
  tableStyleFlags,
  tableActiveCell,
  onEditChart,
  onOpenChartDataDialog,
  onArrange,
  onFlip,
  canDistribute,
}: Props) {
  const { t } = useI18n()
  // Shapes get their own format tab; text-bearing shapes do not auto-activate it
  // (users are usually after Home's text controls when selecting them).
  const contextTab = contextTabForElement(contextElementType ?? null)
  const autoContextTab = autoContextTabForElement(contextElementType ?? null)

  const [tab, setTab] = useState<MainTab | ContextTab>('home')
  const [fileOpen, setFileOpen] = useState(false)
  const [colorOpen, setColorOpen] = useState(false)
  const [fontOpen, setFontOpen] = useState(false)
  const [sizeOpen, setSizeOpen] = useState(false)
  const [lineSpacingOpen, setLineSpacingOpen] = useState(false)
  const [paraOpen, setParaOpen] = useState(false)
  const [layoutPickOpen, setLayoutPickOpen] = useState(false)
  const [slideSizeOpen, setSlideSizeOpen] = useState(false)
  const [transparencyOpen, setTransparencyOpen] = useState(false)
  const [pictureBorderOpen, setPictureBorderOpen] = useState(false)
  const [shapeStyleOpen, setShapeStyleOpen] = useState(false)
  const [shapeFillOpen, setShapeFillOpen] = useState(false)
  // Debounced picture-border commit: color drags fire repeatedly, and a pending
  // color commit must not clobber a width click landing meanwhile
  const pictureBorderTimer = useRef<number | null>(null)
  const pictureBorderDraft = useRef<{ color: string; widthPt: number } | null>(null)
  const commitPictureBorder = (
    patch: Partial<{ color: string; widthPt: number }>,
    immediate = false,
  ) => {
    if (pictureBorderTimer.current) window.clearTimeout(pictureBorderTimer.current)
    const base = pictureBorderDraft.current ?? {
      color: toPickerHex(contextPictureStroke?.color) ?? '#000000',
      widthPt: contextPictureStroke?.widthPt ?? 1,
    }
    const draft = { ...base, ...patch }
    pictureBorderDraft.current = draft
    const fire = () => {
      pictureBorderTimer.current = null
      onPictureStroke?.({
        ...draft,
        ...(contextPictureStroke?.dashPreset ? { dash: contextPictureStroke.dashPreset } : {}),
      })
    }
    if (immediate) fire()
    else pictureBorderTimer.current = window.setTimeout(fire, 200)
  }
  const [lastColor, setLastColor] = useState('#C43E1C')
  // Bullet color "more colors" native picker echo
  const [lastBulletColor, setLastBulletColor] = useState('#C43E1C')
  // Font-size combobox draft: non-null while the input is focused (typed but not yet applied)
  const [sizeDraft, setSizeDraft] = useState<string | null>(null)
  // Font-family combobox draft: free-typed names cover weight variants absent from the list
  const [fontDraft, setFontDraft] = useState<string | null>(null)
  const [tableOpen, setTableOpen] = useState(false)
  const [tableHover, setTableHover] = useState({ r: 0, c: 0 })
  const [tableCustom, setTableCustom] = useState({ r: 8, c: 5 })
  const [layoutOpen, setLayoutOpen] = useState(false)
  // responsive-collapse state (see the collapse effect below)
  const [collapsedGroups, setCollapsedGroups] = useState<string[]>([])
  const [collapseOpen, setCollapseOpen] = useState<string | null>(null)
  const [translateOpen, setTranslateOpen] = useState(false)
  const [arrangeOpen, setArrangeOpen] = useState(false)
  const [slideShowOpen, setSlideShowOpen] = useState(false)
  const [slideShowFromStart, setSlideShowFromStart] = useState(false)
  // Insert tab dropdown galleries (at most one open at a time)
  const [insertDrop, setInsertDrop] = useState<
    'shapes' | 'icons' | 'chart' | 'smartart' | 'wordart' | 'zoom' | 'addanim' | null
  >(null)
  // Chart design: dropdown panels (add chart element / change colors, at most one open at a time)
  const [chartDrop, setChartDrop] = useState<'elements' | 'colors' | null>(null)
  // Draw tab pen gallery: per-preset customisations live for the session;
  // clicking the already-selected pen opens its color/width flyout
  const [penPresets, setPenPresets] = useState<PenPreset[]>(DEFAULT_PEN_PRESETS)
  const [selectedPen, setSelectedPen] = useState(0)
  const [penFlyout, setPenFlyout] = useState<{ index: number; x: number; y: number } | null>(null)
  const [chartTypeDlgOpen, setChartTypeDlgOpen] = useState(false)
  const chartTitleRef = useRef<HTMLInputElement>(null)
  const catAxisRef = useRef<HTMLInputElement>(null)
  const valAxisRef = useRef<HTMLInputElement>(null)
  const [iconColor, setIconColor] = useState(ICON_COLORS[0]!)

  // One ribbon popup at a time: every dropdown trigger closes its siblings on
  // mousedown (via closeSiblingPanels) before its own click-toggle runs, so
  // popups never stack. `keep` names the popups that must survive the sweep —
  // the trigger's own (its toggle decides) and, for nested triggers, the panel
  // anchoring them.
  const closePanels = useCallback((keep: RibbonPanelKey[] = []) => {
    if (!keep.includes('file')) setFileOpen(false)
    if (!keep.includes('color')) setColorOpen(false)
    if (!keep.includes('font')) setFontOpen(false)
    if (!keep.includes('size')) setSizeOpen(false)
    if (!keep.includes('lineSpacing')) setLineSpacingOpen(false)
    if (!keep.includes('para')) setParaOpen(false)
    if (!keep.includes('layoutPick')) setLayoutPickOpen(false)
    if (!keep.includes('slideSize')) setSlideSizeOpen(false)
    if (!keep.includes('transparency')) setTransparencyOpen(false)
    if (!keep.includes('pictureBorder')) setPictureBorderOpen(false)
    if (!keep.includes('shapeStyle')) setShapeStyleOpen(false)
    if (!keep.includes('shapeFill')) setShapeFillOpen(false)
    if (!keep.includes('table')) setTableOpen(false)
    if (!keep.includes('layout')) setLayoutOpen(false)
    if (!keep.includes('translate')) setTranslateOpen(false)
    if (!keep.includes('arrange')) setArrangeOpen(false)
    if (!keep.includes('insert')) setInsertDrop(null)
    if (!keep.includes('chart')) setChartDrop(null)
    if (!keep.includes('collapse')) setCollapseOpen(null)
    if (!keep.includes('pen')) setPenFlyout(null)
    if (!keep.includes('slideShow')) setSlideShowOpen(false)
  }, [])

  // Any ribbon popup open? Drives outside-press dismissal AND suspends the
  // ribbon-tabs window drag region (drag regions swallow mousedown, so a
  // press there could never dismiss otherwise)
  const anyPanelOpen =
    tableOpen ||
    colorOpen ||
    translateOpen ||
    insertDrop != null ||
    fontOpen ||
    sizeOpen ||
    layoutOpen ||
    chartDrop != null ||
    arrangeOpen ||
    slideShowOpen ||
    paraOpen ||
    pictureBorderOpen ||
    shapeStyleOpen ||
    shapeFillOpen ||
    layoutPickOpen ||
    slideSizeOpen ||
    transparencyOpen ||
    lineSpacingOpen ||
    collapseOpen != null

  // Clicking elsewhere collapses every popup (the font color palette uses
  // onMouseDown without stealing focus, collapsing naturally when the edit
  // commits). The shared installer covers outside mousedown, window blur and
  // the shell app:chrome-pressed relay; panels survive via stopPropagation.
  useDismissablePopover(anyPanelOpen, closePanels)

  // ── Responsive collapse (PowerPoint model): the collapsed set is a pure
  // function of the current width, never of resize history — pick the fewest
  // COLLAPSE_ORDER groups whose folding lets the full inline layout fit.
  // Expanded/collapsed widths are cached per group so the required width is
  // computable in every state (before the first fold the collapsed width is
  // an estimate, corrected by measurement as soon as the group first folds).
  const bodyRef = useRef<HTMLDivElement | null>(null)
  const inlineWidthsRef = useRef(new Map<string, number>())
  const collapsedWidthsRef = useRef(new Map<string, number>())
  useLayoutEffect(() => {
    setCollapsedGroups([])
    setCollapseOpen(null)
  }, [tab])
  useLayoutEffect(() => {
    const el = bodyRef.current
    if (!el) return
    const order = COLLAPSE_ORDER[tab] ?? []
    if (!order.length) return
    const evaluate = () => {
      const kids = Array.from(el.children) as HTMLElement[]
      if (!kids.length) return
      const first = kids[0]!
      const last = kids[kids.length - 1]!
      let fullWidth = last.offsetLeft + last.offsetWidth - first.offsetLeft
      // refresh width caches and normalize the measured extent to "all expanded"
      const saving = (g: string) =>
        Math.max(
          0,
          (inlineWidthsRef.current.get(g) ?? 240) - (collapsedWidthsRef.current.get(g) ?? 68),
        )
      for (const g of order) {
        const groupEl = el.querySelector<HTMLElement>(`[data-rbgroup="${g}"]`)
        if (!groupEl) continue
        if (collapsedGroups.includes(g)) {
          collapsedWidthsRef.current.set(g, groupEl.offsetWidth)
          fullWidth += saving(g)
        } else {
          inlineWidthsRef.current.set(g, groupEl.offsetWidth)
        }
      }
      // fewest folded groups whose savings make the layout fit `avail`
      const fitCount = (avail: number) => {
        let need = fullWidth
        let k = 0
        while (k < order.length && need > avail) {
          need -= saving(order[k]!)
          k++
        }
        return k
      }
      const mustCollapse = fitCount(el.clientWidth)
      // integer offset* measurements make the normalized width jitter by a
      // couple of px — demand a little real slack before unfolding so a
      // borderline width can't oscillate
      const next =
        mustCollapse >= collapsedGroups.length
          ? mustCollapse
          : Math.min(Math.max(fitCount(el.clientWidth - 8), mustCollapse), collapsedGroups.length)
      if (next !== collapsedGroups.length) setCollapsedGroups(order.slice(0, next))
    }
    evaluate()
    const ro = new ResizeObserver(evaluate)
    ro.observe(el)
    // group contents can change width without the body resizing (font loads,
    // locale, contextual controls) — watch every group as well
    el.querySelectorAll<HTMLElement>('.ribbon-group').forEach((g) => ro.observe(g))
    return () => ro.disconnect()
  }, [tab, collapsedGroups])

  // Contextual tab auto-switch: jump in for dedicated object tools, but not text-bearing shapes.
  // Track visibility separately so a manually opened picture-format tab still closes on deselect.
  const prevContextTab = useRef<ContextTab | null>(null)
  const prevAutoContextTab = useRef<ContextTab | null>(null)
  useEffect(() => {
    const previousContextTab = prevContextTab.current
    const previousAutoContextTab = prevAutoContextTab.current
    if (autoContextTab && autoContextTab !== previousAutoContextTab) {
      setTab(autoContextTab)
    } else if (!autoContextTab && previousAutoContextTab) {
      setTab((cur) => (cur === previousAutoContextTab ? 'home' : cur))
    } else if (!contextTab && previousContextTab) {
      setTab((cur) => (cur === previousContextTab ? 'home' : cur))
    }
    prevContextTab.current = contextTab
    prevAutoContextTab.current = autoContextTab
  }, [contextTab, autoContextTab])

  /** Insert tab dropdown big button (click toggles, content stopPropagation) */
  const dropBig = (
    key: NonNullable<typeof insertDrop>,
    icon: ReactNode,
    label: string,
    title: string,
    content: ReactNode,
    disabled = !hasDoc,
  ) => (
    <div className="rb-drop-wrap">
      <button
        className={`rb-big ${insertDrop === key ? 'active' : ''}`}
        disabled={disabled}
        data-tip={title}
        onMouseDown={(e) => {
          e.stopPropagation()
          closeSiblingPanels(e, closePanels, 'insert')
        }}
        onClick={() => setInsertDrop((v) => (v === key ? null : key))}
      >
        <span className="rb-big-icon">
          {icon}
          <RbCaret />
        </span>
        <span>{label}</span>
      </button>
      {insertDrop === key && (
        <div className="rb-drop" onMouseDown={(e) => e.stopPropagation()}>
          {content}
        </div>
      )}
    </div>
  )
  // Apply a typed font size: any positive value, 0.5pt steps, clamped to 1-999.
  // While text-editing, restore the selection saved when the input took focus so the size applies
  // to the selection instead of element-level
  const commitSizeDraft = () => {
    const v = parseFloat((sizeDraft ?? '').replace(',', '.'))
    if (!Number.isFinite(v) || v <= 0) return
    const pt = Math.min(999, Math.max(1, Math.round(v * 2) / 2))
    if (editing) restoreEditSelection()
    onFontSize(pt)
  }

  // Apply a free-typed font name, same selection dance as commitSizeDraft
  const commitFontDraft = () => {
    const v = (fontDraft ?? '').trim()
    if (!v) return
    if (editing) restoreEditSelection()
    onFontFamily(v)
  }

  // Custom font color via the native picker: debounced (the picker fires onChange
  // continuously while dragging). The picker steals focus, so while editing the
  // saved selection is restored before each apply
  const customColorTimer = useRef<number | null>(null)
  const onCustomTextColor = (value: string) => {
    const hex = value.toUpperCase()
    setLastColor(hex)
    if (customColorTimer.current) window.clearTimeout(customColorTimer.current)
    customColorTimer.current = window.setTimeout(() => {
      if (editing) {
        restoreEditSelection()
        onTextColor(hex)
      } else onElementTextColor(hex)
    }, 200)
  }

  // Hover preview for animation effects: fire after a short dwell so
  // sweeping across the gallery doesn't spam previews; leaving cancels/stops.
  const animHoverTimer = useRef<number | null>(null)
  const animHoverStart = (effect: AnimEffectKind, motionPath?: string) => {
    if (animHoverTimer.current) window.clearTimeout(animHoverTimer.current)
    animHoverTimer.current = window.setTimeout(() => {
      animHoverTimer.current = null
      onAnimHoverPreview(effect, motionPath)
    }, 350)
  }
  const animHoverStop = () => {
    if (animHoverTimer.current) window.clearTimeout(animHoverTimer.current)
    animHoverTimer.current = null
    onAnimHoverEnd()
  }

  // Custom bullet color via the native picker: same debounce as font color
  const bulletColorTimer = useRef<number | null>(null)
  const onCustomBulletColor = (value: string) => {
    const hex = value.toUpperCase()
    setLastBulletColor(hex)
    if (bulletColorTimer.current) window.clearTimeout(bulletColorTimer.current)
    bulletColorTimer.current = window.setTimeout(() => onParagraphFormat({ bulletColor: hex }), 200)
  }

  // One-time acknowledgement before whole-document AI rewrites:
  // Spelling / Translate send the full deck to the agent, consume credits and may
  // rewrite every slide — say so once before the first run.
  const confirmAiRewrite = () => {
    if (localStorage.getItem(AI_REWRITE_ACK_KEY) === '1') return true
    if (!window.confirm(t('ribbonAiRewriteConfirm'))) return false
    localStorage.setItem(AI_REWRITE_ACK_KEY, '1')
    return true
  }

  // Format buttons use onMouseDown+preventDefault, avoiding stealing contentEditable focus and triggering a commit
  const fmtBtn = (cmd: FormatCmd, label: ReactNode, title: string, className?: string) => (
    <button
      className={`rb-icon${className ? ` ${className}` : ''}`}
      disabled={!editing}
      data-tip={editing ? title : t('ribbonEditableHint', { title })}
      aria-label={title}
      onMouseDown={(e) => {
        e.preventDefault()
        if (editing) onFormat(cmd)
      }}
    >
      {label}
    </button>
  )

  const tabCtx: RibbonTabCtx = {
    aiOpen,
    brushMode,
    canDistribute,
    canPaste,
    closePanels,
    curBulletChar,
    curAlign,
    curFontFamily,
    curFontSizeMixed,
    curFontSizePt,
    currentSlide,
    deckEmpty,
    editing,
    formatOpen,
    hasBrushFormat,
    hasDoc,
    hasSelection,
    hasTextSelection,
    layouts,
    layoutSize,
    onAddSection,
    onAddSlide,
    onAddSlideWithLayout,
    onAiPreset,
    onAlign,
    onArrange,
    onFlip,
    onCopy,
    onCut,
    onElementTextColor,
    onFindReplace,
    onFontFamily,
    onFontSize,
    onFormat,
    onFormatBrushClick,
    onFormatBrushDoubleClick,
    onInsert,
    onPickShape,
    onInsertChart,
    onInsertField,
    onInsertIcon,
    onInsertImage,
    onInsertMedia,
    onInsertModel3d,
    onInsertSmartArt,
    onInsertTable,
    onInsertWordArt,
    onInsertZoom,
    onNewComment,
    onOpenEquation,
    onOpenHeaderFooter,
    onOpenLink,
    onParagraphFormat,
    onPaste,
    onResetLayout,
    onSetLayout,
    onSlideShow,
    onStrike,
    onTextColor,
    onTextToggle,
    onToggleAi,
    onToggleFormat,
    onToggleScreenRecord,
    recording,
    slideCount,
    zoom,
    arrangeOpen,
    collapseOpen,
    collapsedGroups,
    colorOpen,
    commitFontDraft,
    commitSizeDraft,
    dropBig,
    fontDraft,
    setFontDraft,
    fmtBtn,
    fontOpen,
    iconColor,
    lastBulletColor,
    lastColor,
    layoutOpen,
    layoutPickOpen,
    lineSpacingOpen,
    onCustomBulletColor,
    onCustomTextColor,
    paraOpen,
    setArrangeOpen,
    setCollapseOpen,
    setColorOpen,
    setFontOpen,
    setIconColor,
    setInsertDrop,
    setLastColor,
    setLayoutOpen,
    setLayoutPickOpen,
    setLineSpacingOpen,
    setParaOpen,
    setSizeDraft,
    setSizeOpen,
    setSlideShowFromStart,
    setSlideShowOpen,
    setTableCustom,
    setTableHover,
    setTableOpen,
    sizeDraft,
    sizeOpen,
    slideShowFromStart,
    slideShowOpen,
    t,
    tableCustom,
    tableHover,
    tableOpen,
  }

  return (
    <div className="ribbon">
      <div
        className={`ribbon-tabs ${IN_TAB ? '' : IS_MAC ? 'ribbon-tabs-mac' : 'ribbon-tabs-win'}${
          anyPanelOpen ? ' ribbon-tabs-nodrag' : ''
        }`}
      >
        {!IS_MAC && (
          <div className="file-tab-wrap">
            <button
              className={`ribbon-tab ribbon-tab-file ${fileOpen ? 'open' : ''}`}
              onMouseDown={(e) => {
                e.stopPropagation()
                closeSiblingPanels(e, closePanels, 'file')
              }}
              onClick={() => setFileOpen((v) => !v)}
            >
              {t('ribbonTabFile')}
            </button>
            {fileOpen && (
              <div className="file-menu">
                <button
                  onClick={() => {
                    setFileOpen(false)
                    onOpen()
                  }}
                >
                  {t('ribbonFileOpen')} <span className="file-menu-key">Ctrl+O</span>
                </button>
                <button
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onSave()
                  }}
                >
                  {t('ribbonFileSave')} <span className="file-menu-key">Ctrl+S</span>
                </button>
                <button
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onSaveAs()
                  }}
                >
                  {t('ribbonFileSaveAs')} <span className="file-menu-key">Ctrl+Shift+S</span>
                </button>
                <button
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onExportPdf()
                  }}
                >
                  {t('ribbonFileExportPdf')}
                </button>
                <button
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onPrint()
                  }}
                >
                  {t('ribbonFilePrint')} <span className="file-menu-key">Ctrl+P</span>
                </button>
                <button
                  disabled={!hasDoc}
                  onClick={() => {
                    setFileOpen(false)
                    onExportImages()
                  }}
                >
                  {t('ribbonFileExportImages')}
                </button>
              </div>
            )}
          </div>
        )}
        <button
          className="qa-btn"
          data-tip={t('ribbonSaveTip')}
          aria-label={t('ribbonSaveTip')}
          disabled={!dirty}
          onClick={onSave}
        >
          <IconSave size={16} />
        </button>
        {/* onMouseDown+preventDefault like the format buttons: keep contentEditable focus so undo/redo reaches
            the active text edit. onClick with detail===0 covers keyboard activation (Enter/Space emit only click). */}
        <button
          className="qa-btn"
          data-tip={t('ribbonUndo')}
          aria-label={t('ribbonUndo')}
          disabled={!hasDoc || (!canUndo && !editing)}
          onMouseDown={(e) => {
            e.preventDefault()
            onUndo()
          }}
          onClick={(e) => {
            if (e.detail === 0) onUndo()
          }}
        >
          <IconUndo size={16} />
        </button>
        <button
          className="qa-btn"
          data-tip={t('ribbonRedo')}
          aria-label={t('ribbonRedo')}
          disabled={!hasDoc || (!canRedo && !editing)}
          onMouseDown={(e) => {
            e.preventDefault()
            onRedo()
          }}
          onClick={(e) => {
            if (e.detail === 0) onRedo()
          }}
        >
          <IconRedo size={16} />
        </button>
        <label
          className={`autosave-toggle ${autoSave ? 'on' : ''}`}
          data-tip={t('ribbonAutoSaveTip')}
        >
          <span className="autosave-knob" />
          <span className="autosave-text">{t('ribbonAutoSave')}</span>
          <input
            type="checkbox"
            checked={autoSave}
            onChange={(e) => onAutoSaveChange(e.target.checked)}
          />
        </label>
        <span className="qa-sep" aria-hidden="true" />
        {TABS.filter((tb) => tb !== 'file').map((tb) => (
          <button
            key={tb}
            className={`ribbon-tab ${tab === tb ? 'active' : ''}`}
            onClick={() => {
              setTab(tb)
              setFileOpen(false)
            }}
          >
            {t(TAB_LABEL[tb])}
          </button>
        ))}
        {contextTab && (
          <button
            key={contextTab}
            className={`ribbon-tab ribbon-tab-context ${tab === contextTab ? 'active' : ''}`}
            onClick={() => setTab(contextTab)}
            data-tip={t(TAB_LABEL[contextTab])}
          >
            {t(TAB_LABEL[contextTab])}
          </button>
        )}
        <span className="ribbon-tabs-spacer" />
      </div>

      <div className="ribbon-body" ref={bodyRef}>
        {tab === 'home' ? (
          <RibbonHomeTab rb={tabCtx} />
        ) : tab === 'insert' ? (
          <RibbonInsertTab rb={tabCtx} />
        ) : tab === 'draw' ? (
          (() => {
            // Draw tab: tool buttons, then a tray of ready pens.
            // Clicking a pen picks it up; clicking the held pen opens its
            // color/width flyout; customisations stick to that pen preset.
            const applyPenPreset = (preset: PenPreset) => {
              // onInkTool toggles back to 'select' when the active tool is re-picked;
              // only switch when it actually differs so editing color/width (or swapping
              // to another pen of the same kind) never drops the pen
              if (inkTool !== preset.kind) onInkTool(preset.kind)
              if (preset.kind === 'pen')
                onInkPen({ ...inkPen, color: preset.color, width: preset.width })
              else onInkHighlighter({ ...inkHighlighter, color: preset.color, width: preset.width })
            }
            const updatePenPreset = (
              index: number,
              patch: Partial<Pick<PenPreset, 'color' | 'width'>>,
            ) => {
              const next = penPresets.map((p, i) => (i === index ? { ...p, ...patch } : p))
              setPenPresets(next)
              applyPenPreset(next[index])
            }
            // Color swatches + width dots editing one pen preset (same controls as
            // apps/docs DrawTab); shown inline for the selected pen and reused by the
            // held-pen flyout
            const renderInkSettings = (index: number) => {
              const preset = penPresets[index]!
              const widths = preset.kind === 'highlighter' ? HIGHLIGHTER_WIDTHS : PEN_WIDTHS
              return (
                <div className="ink-settings">
                  <div className="ink-swatches">
                    {INK_COLORS.map((hex) => (
                      <button
                        key={hex}
                        className={`ink-swatch ${preset.color === hex ? 'active' : ''}`}
                        style={{ background: `#${hex}` }}
                        data-tip={`#${hex}`}
                        aria-label={`#${hex}`}
                        disabled={!hasDoc}
                        onClick={() => updatePenPreset(index, { color: hex })}
                      />
                    ))}
                  </div>
                  <div className="ink-widths">
                    {widths.map((w) => (
                      <button
                        key={w}
                        className={`ink-width ${preset.width === w ? 'active' : ''}`}
                        data-tip={t('ribbonInkWidthTip', { w })}
                        aria-label={t('ribbonInkWidthTip', { w })}
                        disabled={!hasDoc}
                        onClick={() => updatePenPreset(index, { width: w })}
                      >
                        <span
                          className="ink-width-dot"
                          style={{
                            width: Math.min(16, w * 2 + 2),
                            height: Math.min(16, w * 2 + 2),
                            background: `#${preset.color}`,
                          }}
                        />
                      </button>
                    ))}
                  </div>
                </div>
              )
            }
            return (
              <>
                <Group label={t('ribbonGroupDrawTools')}>
                  <button
                    className={`rb-big ${inkTool === 'select' ? 'active' : ''}`}
                    disabled={!hasDoc}
                    data-tip={t('ribbonSelectTip')}
                    onClick={() => {
                      setPenFlyout(null)
                      onInkTool('select')
                    }}
                  >
                    <span className="rb-big-icon">
                      <IconCursor size={BIG} />
                    </span>
                    <span>{t('ribbonGroupSelect')}</span>
                  </button>
                  <button
                    className={`rb-big ${inkTool === 'eraser' ? 'active' : ''}`}
                    disabled={!hasDoc}
                    data-tip={t('ribbonEraserTip')}
                    onClick={() => {
                      setPenFlyout(null)
                      onInkTool('eraser')
                    }}
                  >
                    <span className="rb-big-icon">
                      <IconEraser size={BIG} />
                    </span>
                    <span>{t('ribbonEraser')}</span>
                  </button>
                </Group>
                <div className="ribbon-sep" />
                <Group label={t('ribbonGroupPenStyle')}>
                  <div className="pen-tray">
                    {penPresets.map((preset, i) => {
                      const held = selectedPen === i && inkTool === preset.kind
                      return (
                        <button
                          key={i}
                          className={`pen-btn ${held ? 'down' : ''}`}
                          disabled={!hasDoc}
                          data-tip={
                            preset.kind === 'pen' ? t('ribbonPenTip') : t('ribbonHighlighterTip')
                          }
                          aria-label={
                            preset.kind === 'pen' ? t('ribbonPenTip') : t('ribbonHighlighterTip')
                          }
                          onClick={(event) => {
                            if (held) {
                              const rect = event.currentTarget.getBoundingClientRect()
                              setPenFlyout(
                                penFlyout?.index === i
                                  ? null
                                  : {
                                      index: i,
                                      x: Math.round(rect.left),
                                      y: Math.round(rect.bottom + 6),
                                    },
                              )
                            } else {
                              setSelectedPen(i)
                              setPenFlyout(null)
                              applyPenPreset(preset)
                            }
                          }}
                        >
                          <PenThumb kind={preset.kind} color={preset.color} />
                        </button>
                      )
                    })}
                  </div>
                  {renderInkSettings(selectedPen)}
                </Group>
                <div className="ribbon-sep" />
                <Group label={t('ribbonGroupClear')}>
                  <button
                    className="rb-big"
                    disabled={!hasDoc || inkCount === 0}
                    data-tip={t('ribbonEraseAllTip')}
                    onClick={onInkClearAll}
                  >
                    <span className="rb-big-icon">
                      <IconEraser size={BIG} />
                    </span>
                    <span>{t('ribbonEraseAll')}</span>
                  </button>
                </Group>
                {penFlyout && (
                  <>
                    <div className="pen-flyout-backdrop" onClick={() => setPenFlyout(null)} />
                    <div className="pen-flyout" style={{ left: penFlyout.x, top: penFlyout.y }}>
                      {renderInkSettings(penFlyout.index)}
                    </div>
                  </>
                )}
              </>
            )
          })()
        ) : tab === 'design' ? (
          <>
            <Group label={t('ribbonGroupThemes')}>
              <div className="theme-gallery">
                {THEME_PRESETS.map((tp) => (
                  <button
                    key={tp.id}
                    className="theme-card"
                    disabled={!hasDoc}
                    data-tip={t('ribbonApplyThemeTip', { name: themeDisplayName(tp, t) })}
                    onClick={() => onApplyTheme(tp)}
                    style={{ background: `#${tp.colors.lt1}`, color: `#${tp.colors.dk1}` }}
                  >
                    <span className="theme-card-aa" style={{ fontFamily: tp.majorFont }}>
                      Aa
                    </span>
                    <span className="theme-card-dots">
                      {['accent1', 'accent2', 'accent3', 'accent4'].map((k) => (
                        <span
                          key={k}
                          className="theme-card-dot"
                          style={{ background: `#${tp.colors[k]}` }}
                        />
                      ))}
                    </span>
                    <span className="theme-card-name">{themeDisplayName(tp, t)}</span>
                  </button>
                ))}
              </div>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupBackground')}>
              <button
                className="rb-big"
                disabled={!hasDoc}
                onClick={onFormatBackground}
                data-tip={t('ribbonFormatBackgroundTip')}
              >
                <span className="rb-big-icon">
                  <IconPageColor size={BIG} />
                </span>
                <span>{t('ribbonFormatBackground')}</span>
              </button>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupCustomize')}>
              <div className="rb-drop-wrap">
                <button
                  className={`rb-big ${slideSizeOpen ? 'active' : ''}`}
                  disabled={!hasDoc}
                  onMouseDown={(e) => {
                    e.stopPropagation()
                    closeSiblingPanels(e, closePanels, 'slideSize')
                  }}
                  onClick={() => setSlideSizeOpen((v) => !v)}
                  data-tip={t('ribbonSlideSizeTip')}
                >
                  <span className="rb-big-icon">
                    <IconPageSize size={BIG} />
                    <RbCaret />
                  </span>
                  <span>{t('ribbonSlideSize')}</span>
                </button>
                {slideSizeOpen && (
                  <div className="rb-drop rb-menu" onMouseDown={(e) => e.stopPropagation()}>
                    {(
                      [
                        ['16:9', t('ribbonSlideSize169'), 12192000, 6858000],
                        ['4:3', t('ribbonSlideSize43'), 9144000, 6858000],
                      ] as const
                    ).map(([key, label, cx, cy]) => (
                      <button
                        key={key}
                        className={slideSizeKey === key ? 'on' : ''}
                        onClick={() => {
                          setSlideSizeOpen(false)
                          onSlideSize(cx, cy)
                        }}
                      >
                        {label}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            </Group>
          </>
        ) : tab === 'transitions' ? (
          <>
            <Group label={t('ribbonGroupTransitionToThis')}>
              {TRANSITIONS.map((tr) => (
                <button
                  key={tr.kind}
                  className={`rb-big ${transition === tr.kind ? 'active' : ''}`}
                  disabled={!hasDoc}
                  onClick={() => onTransition(tr.kind, false)}
                  data-tip={
                    tr.kind === 'none'
                      ? t('ribbonTransNoneTip')
                      : t('ribbonTransApplyTip', { name: t(tr.label) })
                  }
                >
                  <span className="rb-big-icon rb-trans-glyph">{tr.icon}</span>
                  <span>{t(tr.label)}</span>
                </button>
              ))}
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupTiming')}>
              <button
                className="rb-big"
                disabled={!hasDoc}
                onClick={() => onTransition(transition, true)}
                data-tip={t('ribbonTransApplyAllTip')}
              >
                <span className="rb-big-icon">
                  <IconApplyAll size={BIG} />
                </span>
                <span>{t('ribbonApplyToAll')}</span>
              </button>
            </Group>
          </>
        ) : tab === 'animations' ? (
          <>
            <Group label={t('ribbonPreview')}>
              <button
                className="rb-big"
                disabled={!hasDoc || animCount === 0}
                onClick={onAnimPreview}
                data-tip={t('ribbonAnimPreviewTip')}
              >
                <span className="rb-big-icon">
                  <IconPlayCurrent size={BIG} />
                </span>
                <span>{t('ribbonPreview')}</span>
              </button>
            </Group>
            <div className="ribbon-sep" />
            <Group
              label={t('ribbonGroupAnimation')}
              groupId="animation"
              collapse={{
                collapsed: collapsedGroups.includes('animation'),
                open: collapseOpen === 'animation',
                onToggle: () => {
                  closePanels(['collapse'])
                  setCollapseOpen((v) => (v === 'animation' ? null : 'animation'))
                },
                icon: (
                  <span className="rb-anim-glyph rb-anim-entr">
                    <IconAnimStar size={20} />
                  </span>
                ),
              }}
            >
              <button
                className="rb-big"
                disabled={!hasDoc || !hasSelection}
                onClick={() => onApplyAnimation('none')}
                data-tip={t('ribbonAnimNoneTip')}
              >
                <span className="rb-big-icon rb-anim-glyph">
                  <IconAnimNone size={BIG} />
                </span>
                <span>{t('ribbonNone')}</span>
              </button>
              {ANIM_EFFECTS.map((a) => (
                <button
                  key={a.kind}
                  className={`rb-big ${selectedAnimEffect === a.kind ? 'active' : ''}`}
                  disabled={!hasDoc || !hasSelection}
                  onClick={() => onApplyAnimation(a.kind)}
                  onMouseEnter={() => {
                    if (hasDoc && hasSelection) animHoverStart(a.kind)
                  }}
                  onMouseLeave={animHoverStop}
                  data-tip={t('ribbonAnimApplyTip', {
                    cls: t(ANIM_CLS_TITLE[a.cls]),
                    name: t(a.label),
                  })}
                >
                  <span className={`rb-big-icon rb-anim-glyph rb-anim-${a.cls}`}>
                    <AnimEffectIcon kind={a.kind} size={BIG} />
                  </span>
                  <span>{t(a.label)}</span>
                </button>
              ))}
            </Group>
            <div className="ribbon-sep" />
            <Group
              label={t('ribbonGroupMotionPaths')}
              groupId="motionPaths"
              collapse={{
                collapsed: collapsedGroups.includes('motionPaths'),
                open: collapseOpen === 'motionPaths',
                onToggle: () => {
                  closePanels(['collapse'])
                  setCollapseOpen((v) => (v === 'motionPaths' ? null : 'motionPaths'))
                },
                icon: (
                  <span className="rb-anim-glyph rb-anim-path">
                    <IconPathDiagonal size={20} />
                  </span>
                ),
              }}
            >
              {MOTION_PATHS.map((mp) => (
                <button
                  key={mp.label}
                  className="rb-big"
                  disabled={!hasDoc || !hasSelection}
                  onClick={() => onApplyMotionPath(mp.path)}
                  onMouseEnter={() => {
                    if (hasDoc && hasSelection) animHoverStart('motionPath', mp.path)
                  }}
                  onMouseLeave={animHoverStop}
                  data-tip={t('ribbonMotionPathTip', { name: t(mp.label) })}
                >
                  <span className="rb-big-icon rb-anim-glyph rb-anim-path">{mp.icon}</span>
                  <span>{t(mp.label)}</span>
                </button>
              ))}
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupAdvancedAnim')}>
              {dropBig(
                'addanim',
                <IconSparkle size={BIG} />,
                t('ribbonAddAnimation'),
                t('ribbonAddAnimationTip'),
                <div className="rb-menu rb-anim-menu">
                  {(['entr', 'emph', 'exit'] as const).map((cls) => (
                    <React.Fragment key={cls}>
                      <div className="rb-drop-title">{t(ANIM_CLS_TITLE[cls])}</div>
                      {ANIM_EFFECTS.filter((a) => a.cls === cls).map((a) => (
                        <button
                          key={a.kind}
                          onClick={() => {
                            onAddAnimation(a.kind)
                            setInsertDrop(null)
                          }}
                          onMouseEnter={() => {
                            if (hasDoc && hasSelection) animHoverStart(a.kind)
                          }}
                          onMouseLeave={animHoverStop}
                        >
                          <span className={`rb-anim-glyph rb-anim-${a.cls}`}>
                            <AnimEffectIcon kind={a.kind} size={15} />
                          </span>{' '}
                          {t(a.label)}
                        </button>
                      ))}
                    </React.Fragment>
                  ))}
                </div>,
                !hasDoc || !hasSelection,
              )}
              <button
                className={`rb-big ${animPaneOpen ? 'active' : ''}`}
                disabled={!hasDoc}
                onClick={onToggleAnimPane}
                data-tip={t('ribbonAnimPaneTip')}
              >
                <span className="rb-big-icon">
                  <IconNavPane size={BIG} />
                </span>
                <span>{t('ribbonAnimPane')}</span>
              </button>
              <button
                className={`rb-big ${animByParagraph ? 'active' : ''}`}
                disabled={!hasDoc}
                onClick={onToggleAnimByParagraph}
                data-tip={t('ribbonAnimByParaTip')}
              >
                <span className="rb-big-icon">
                  <IconBullets size={BIG} />
                </span>
                <span>{t('ribbonAnimByPara')}</span>
              </button>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupTiming')}>
              <div className="rb-anim-timing">
                <label>
                  {t('ribbonAnimStart')}
                  <select
                    disabled={!timingAnim}
                    value={timingAnim?.trigger ?? 'onClick'}
                    onChange={(e) => onAnimTiming({ trigger: e.target.value as AnimTrigger })}
                    data-tip={t('ribbonAnimTriggerTip')}
                  >
                    <option value="onClick">{t('ribbonAnimOnClick')}</option>
                    <option value="withPrev">{t('ribbonAnimWithPrev')}</option>
                    <option value="afterPrev">{t('ribbonAnimAfterPrev')}</option>
                  </select>
                </label>
                <label>
                  {t('ribbonAnimDuration')}
                  <input
                    key={`dur-${timingAnim?.sourceId ?? ''}-${timingAnim?.durationMs ?? ''}`}
                    type="number"
                    min={0}
                    step={0.1}
                    disabled={!timingAnim}
                    defaultValue={timingAnim ? (timingAnim.durationMs / 1000).toFixed(2) : ''}
                    onBlur={(e) => {
                      const v = parseFloat(e.target.value)
                      if (
                        timingAnim &&
                        Number.isFinite(v) &&
                        Math.round(v * 1000) !== timingAnim.durationMs
                      )
                        onAnimTiming({ durationMs: Math.max(0, Math.round(v * 1000)) })
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                    }}
                    title={t('ribbonAnimDurationTip')}
                  />
                  {t('ribbonSecondsUnit')}
                </label>
                <label>
                  {t('ribbonAnimDelay')}
                  <input
                    key={`delay-${timingAnim?.sourceId ?? ''}-${timingAnim?.delayMs ?? ''}`}
                    type="number"
                    min={0}
                    step={0.1}
                    disabled={!timingAnim}
                    defaultValue={timingAnim ? (timingAnim.delayMs / 1000).toFixed(2) : ''}
                    onBlur={(e) => {
                      const v = parseFloat(e.target.value)
                      if (
                        timingAnim &&
                        Number.isFinite(v) &&
                        Math.round(v * 1000) !== timingAnim.delayMs
                      )
                        onAnimTiming({ delayMs: Math.max(0, Math.round(v * 1000)) })
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                    }}
                    title={t('ribbonAnimDelayTip')}
                  />
                  {t('ribbonSecondsUnit')}
                </label>
              </div>
            </Group>
          </>
        ) : tab === 'slideShow' ? (
          <>
            <Group label={t('ribbonGroupStartShow')}>
              <button
                className="rb-big"
                disabled={!hasDoc}
                onClick={() => onSlideShow(true)}
                data-tip={t('ribbonFromBeginningTip')}
              >
                <span className="rb-big-icon">
                  <IconPlayFromStart size={BIG} />
                </span>
                <span>{t('ribbonFromBeginning')}</span>
              </button>
              <button
                className="rb-big"
                disabled={!hasDoc}
                onClick={() => onSlideShow(false)}
                data-tip={t('ribbonFromCurrentTip')}
              >
                <span className="rb-big-icon">
                  <IconPlayCurrent size={BIG} />
                </span>
                <span>{t('ribbonFromCurrent')}</span>
              </button>
              <button
                className="rb-big"
                disabled={!hasDoc}
                onClick={() => onPresenterView(true)}
                data-tip={t('ribbonPresenterViewTip')}
              >
                <span className="rb-big-icon">
                  <IconPresenterView size={BIG} />
                </span>
                <span>{t('ribbonPresenterView')}</span>
              </button>
              <button
                className="rb-big"
                disabled={!hasDoc}
                onClick={onCustomShow}
                data-tip={t('ribbonCustomShowTip')}
              >
                <span className="rb-big-icon">
                  <IconCustomShow size={BIG} />
                </span>
                <span>{t('ribbonCustomShow')}</span>
              </button>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupSetUp')}>
              <DisabledBig icon={<IconSetupShow size={BIG} />} label={t('ribbonSetUpShow')} />
              <button
                className={`rb-big ${currentHidden ? 'active' : ''}`}
                disabled={!hasDoc}
                onClick={onToggleHidden}
                data-tip={currentHidden ? t('ribbonUnhideSlideTip') : t('ribbonHideSlideTip')}
              >
                <span className="rb-big-icon">
                  <IconHideSlide size={BIG} />
                </span>
                <span>{t('ribbonHideSlide')}</span>
              </button>
              <button
                className="rb-big"
                disabled={!hasDoc}
                onClick={onRehearse}
                data-tip={t('ribbonRehearseTip')}
              >
                <span className="rb-big-icon">
                  <IconRehearse size={BIG} />
                </span>
                <span>{t('ribbonRehearse')}</span>
              </button>
              <DisabledBig icon={<IconRecord size={BIG} />} label={t('ribbonRecord')} />
            </Group>
          </>
        ) : tab === 'review' ? (
          <>
            <Group label={t('ribbonGroupProofing')}>
              <button
                className="rb-big"
                disabled={!hasDoc}
                data-tip={`${t('ribbonSpellCheckTip')} — ${t('ribbonAiCreditNote')}`}
                onClick={() => {
                  if (confirmAiRewrite()) onAiPreset(t('ribbonSpellCheckPrompt'))
                }}
              >
                <span className="rb-big-icon">
                  <span className="ai-feature-icon" aria-hidden="true">
                    <img src={iconSpelling} width={22} height={22} alt="" />
                  </span>
                </span>
                <span>{t('ribbonSpellCheck')}</span>
              </button>
              <div className="rb-drop-wrap">
                <button
                  className={`rb-big ${translateOpen ? 'active' : ''}`}
                  disabled={!hasDoc}
                  data-tip={`${t('ribbonTranslateTip')} — ${t('ribbonAiCreditNote')}`}
                  onMouseDown={(e) => {
                    e.stopPropagation()
                    closeSiblingPanels(e, closePanels, 'translate')
                  }}
                  onClick={() => setTranslateOpen((v) => !v)}
                >
                  <span className="rb-big-icon">
                    <span className="ai-feature-icon" aria-hidden="true">
                      <img src={iconTranslate} width={22} height={22} alt="" />
                    </span>
                    <RbCaret />
                  </span>
                  <span>{t('ribbonTranslate')}</span>
                </button>
                {translateOpen && (
                  <div className="rb-drop rb-menu" onMouseDown={(e) => e.stopPropagation()}>
                    {TRANSLATE_TARGETS.map((lang) => (
                      <button
                        key={lang}
                        onClick={() => {
                          setTranslateOpen(false)
                          if (confirmAiRewrite()) {
                            onAiPreset(t('ribbonTranslatePrompt', { lang: t(lang) }))
                          }
                        }}
                      >
                        {t(lang)}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupComments')}>
              <button
                className="rb-big"
                disabled={!hasDoc}
                onClick={onNewComment}
                data-tip={t('ribbonNewCommentTip')}
              >
                <span className="rb-big-icon">
                  <IconComment size={BIG} />
                </span>
                <span>{t('ribbonNewComment')}</span>
              </button>
              <button
                className={`rb-big ${commentsOpen ? 'active' : ''}`}
                disabled={!hasDoc}
                onClick={onToggleComments}
                data-tip={t('ribbonCommentsPaneTip')}
              >
                <span className="rb-big-icon">
                  <IconNavPane size={BIG} />
                </span>
                <span>
                  {t('ribbonCommentsPane')}
                  {commentCount > 0 ? t('ribbonCountSuffix', { n: commentCount }) : ''}
                </span>
              </button>
            </Group>
          </>
        ) : tab === 'view' ? (
          <>
            <Group label={t('ribbonGroupPresentationViews')}>
              {(
                [
                  [
                    'normal',
                    <IconPrintLayout key="n" size={BIG} />,
                    t('ribbonViewNormal'),
                    t('ribbonViewNormalTip'),
                  ],
                  [
                    'outline',
                    <IconOutlineView key="o" size={BIG} />,
                    t('ribbonViewOutline'),
                    t('ribbonViewOutlineTip'),
                  ],
                  [
                    'sorter',
                    <IconArrangeAll key="s" size={BIG} />,
                    t('ribbonViewSorter'),
                    t('ribbonViewSorterTip'),
                  ],
                  [
                    'reading',
                    <IconReadMode key="r" size={BIG} />,
                    t('ribbonViewReading'),
                    t('ribbonViewReadingTip'),
                  ],
                ] as const
              ).map(([mode, icon, label, title]) => (
                <button
                  key={mode}
                  className={`rb-big ${viewMode === mode ? 'active' : ''}`}
                  disabled={!hasDoc}
                  onClick={() => onViewMode(mode)}
                  data-tip={title}
                >
                  <span className="rb-big-icon">{icon}</span>
                  <span>{label}</span>
                </button>
              ))}
              <button
                className="rb-big"
                disabled={!hasDoc}
                onClick={onSlideMaster}
                data-tip={t('ribbonViewMasterTip')}
              >
                <span className="rb-big-icon">
                  <IconSlideMaster size={BIG} />
                </span>
                <span>{t('ribbonViewMaster')}</span>
              </button>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupShow')}>
              <div className="rb-check-grid">
                <RbCheck
                  label={t('ribbonRuler')}
                  on={showRuler}
                  disabled={!hasDoc}
                  title={t('ribbonRulerTip')}
                  onClick={onToggleRuler}
                />
                <RbCheck
                  label={t('ribbonGridlines')}
                  on={showGrid}
                  disabled={!hasDoc}
                  title={t('ribbonGridlinesTip')}
                  onClick={onToggleGrid}
                />
                <RbCheck
                  label={t('ribbonGuides')}
                  on={showGuides}
                  disabled={!hasDoc}
                  title={t('ribbonGuidesTip')}
                  onClick={onToggleGuides}
                />
                <RbCheck
                  label={t('ribbonNotes')}
                  on={showNotes}
                  disabled={!hasDoc}
                  title={t('ribbonNotesTip')}
                  onClick={onToggleNotes}
                />
                <RbCheck
                  label={t('ribbonThumbnailPane')}
                  on={showThumbs}
                  title={t('ribbonThumbnailPaneTip')}
                  onClick={onToggleThumbs}
                />
              </div>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupZoom')}>
              <div className="rb-col">
                <button
                  className="rb-small"
                  disabled={!hasDoc}
                  onClick={() => onZoom((z) => Math.min(z * 1.15, 3))}
                >
                  <IconZoomIn size={18} />
                  <span>{t('ribbonZoomIn')}</span>
                </button>
                <button
                  className="rb-small"
                  disabled={!hasDoc}
                  onClick={() => onZoom((z) => Math.max(z / 1.15, 0.25))}
                >
                  <IconZoomOut size={18} />
                  <span>{t('ribbonZoomOut')}</span>
                </button>
                <button className="rb-small" disabled={!hasDoc} onClick={() => onZoom(1)}>
                  <IconZoom100 size={18} />
                  <span>100%</span>
                </button>
              </div>
              <button
                className="rb-big"
                disabled={!hasDoc}
                onClick={onZoomFit}
                data-tip={t('ribbonFitWindowTip')}
              >
                <span className="rb-big-icon">
                  <IconFitWindow size={BIG} />
                </span>
                <span>{t('ribbonFitWindow')}</span>
              </button>
            </Group>
          </>
        ) : tab === 'tableDesign' ? (
          <>
            <Group label={t('ribbonGroupTableStyles')}>
              {TABLE_STYLE_PRESETS_UI.map((p) => (
                <button
                  key={p.key}
                  className="rb-table-style-card"
                  data-tip={t(p.label)}
                  disabled={!onEditTableStyle}
                  onClick={() => onEditTableStyle?.({ styleName: p.key })}
                >
                  <span className="rb-table-style-preview">
                    <TableMiniPreview {...p.preview} />
                  </span>
                  <span className="rb-table-style-label">{t(p.label)}</span>
                </button>
              ))}
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupTableOptions')}>
              <TableToggleBtn
                label={t('ribbonHeaderRow')}
                on={tableStyleFlags?.firstRow ?? false}
                disabled={!onEditTableStyle}
                onClick={() => onEditTableStyle?.({ firstRow: true })}
                offClick={() => onEditTableStyle?.({ firstRow: false })}
              />
              <TableToggleBtn
                label={t('ribbonBandedRows')}
                on={tableStyleFlags?.bandRow ?? false}
                disabled={!onEditTableStyle}
                onClick={() => onEditTableStyle?.({ bandRow: true })}
                offClick={() => onEditTableStyle?.({ bandRow: false })}
              />
            </Group>
            <div className="ribbon-sep" />
            <Group label={tableActiveCell ? t('ribbonGroupShadingCell') : t('ribbonGroupShading')}>
              <div className="rb-table-shading">
                {TABLE_SHADING_COLORS.map((c) => (
                  <button
                    key={c}
                    className="rb-color-swatch"
                    style={{ background: c }}
                    data-tip={c}
                    aria-label={c}
                    disabled={!onEditTableStyle}
                    // preventDefault keeps a cell text-edit session alive so shading targets that cell
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() =>
                      onEditTableStyle?.({
                        shadingColor: c,
                        ...(tableActiveCell ? { cells: [tableActiveCell] } : {}),
                      })
                    }
                  />
                ))}
                <button
                  className="rb-color-swatch rb-color-none"
                  data-tip={t('ribbonNoShading')}
                  aria-label={t('ribbonNoShading')}
                  disabled={!onEditTableStyle}
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() =>
                    onEditTableStyle?.({
                      shadingColor: 'none',
                      ...(tableActiveCell ? { cells: [tableActiveCell] } : {}),
                    })
                  }
                >
                  <IconNoneX size={12} />
                </button>
              </div>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupBorders')}>
              <div className="rb-table-border-row">
                <button
                  className="rb-icon"
                  data-tip={t('ribbonAllBordersTip')}
                  aria-label={t('ribbonAllBordersTip')}
                  disabled={!onEditTableStyle}
                  onClick={() =>
                    onEditTableStyle?.({
                      borderPreset: 'all',
                      borderColor: '#000000',
                      borderWidthPt: 1,
                    })
                  }
                >
                  ⊞
                </button>
                <button
                  className="rb-icon"
                  data-tip={t('ribbonClearBordersTip')}
                  aria-label={t('ribbonClearBordersTip')}
                  disabled={!onEditTableStyle}
                  onClick={() => onEditTableStyle?.({ borderPreset: 'none' })}
                >
                  ⊟
                </button>
              </div>
              <div className="rb-table-border-row">
                <span className="rb-label">{t('ribbonBorderColorLabel')}</span>
                <input
                  type="color"
                  defaultValue="#000000"
                  className="rb-color-input"
                  data-tip={t('ribbonBorderColorTip')}
                  onPointerDown={(e) => armColorInput(e.currentTarget)}
                  onChange={(e) => onEditTableStyle?.({ borderColor: e.target.value })}
                />
                <span className="rb-label">{t('ribbonBorderWeightLabel')}</span>
                <select
                  className="rb-select-sm"
                  defaultValue="1"
                  data-tip={t('ribbonBorderWeightTip')}
                  onChange={(e) => onEditTableStyle?.({ borderWidthPt: Number(e.target.value) })}
                >
                  <option value="0.5">0.5pt</option>
                  <option value="1">1pt</option>
                  <option value="1.5">1.5pt</option>
                  <option value="2.25">2.25pt</option>
                  <option value="3">3pt</option>
                </select>
              </div>
            </Group>
          </>
        ) : tab === 'chartDesign' ? (
          <>
            {/* Chart Design group order: chart layout | chart styles | data | type */}
            <Group label={t('ribbonGroupChartLayouts')}>
              <div className="rb-drop-wrap">
                <button
                  className={`rb-big ${chartDrop === 'elements' ? 'active' : ''}`}
                  disabled={!onEditChart}
                  data-tip={t('ribbonAddChartElementTip')}
                  onMouseDown={(e) => {
                    e.stopPropagation()
                    closeSiblingPanels(e, closePanels, 'chart')
                  }}
                  onClick={() => setChartDrop((v) => (v === 'elements' ? null : 'elements'))}
                >
                  <span className="rb-big-icon" style={{ fontSize: 20 }}>
                    ➕<RbCaret />
                  </span>
                  <span>{t('ribbonAddChartElement')}</span>
                </button>
                {chartDrop === 'elements' && (
                  <div
                    className="rb-drop rb-chart-elem-drop"
                    onMouseDown={(e) => e.stopPropagation()}
                  >
                    <div className="rb-drop-title">{t('ribbonChartTitle')}</div>
                    <div className="rb-row">
                      <input
                        ref={chartTitleRef}
                        className="rb-input-sm"
                        defaultValue={contextChartStyle?.title ?? ''}
                        placeholder={t('ribbonEmptyHidden')}
                      />
                      <button
                        className="rb-icon"
                        onClick={() =>
                          onEditChart?.({ title: chartTitleRef.current?.value.trim() ?? '' })
                        }
                      >
                        {t('ribbonApply')}
                      </button>
                    </div>
                    <div className="rb-drop-title">{t('ribbonAxisTitles')}</div>
                    <label className="rb-row">
                      <span className="rb-label">{t('ribbonCatAxis')}</span>
                      <input
                        ref={catAxisRef}
                        className="rb-input-sm"
                        defaultValue={contextChartStyle?.catAxisTitle ?? ''}
                        placeholder={t('ribbonEmptyHidden')}
                      />
                    </label>
                    <label className="rb-row">
                      <span className="rb-label">{t('ribbonValAxis')}</span>
                      <input
                        ref={valAxisRef}
                        className="rb-input-sm"
                        defaultValue={contextChartStyle?.valAxisTitle ?? ''}
                        placeholder={t('ribbonEmptyHidden')}
                      />
                    </label>
                    <div className="rb-row">
                      <button
                        className="rb-icon"
                        onClick={() =>
                          onEditChart?.({
                            catAxisTitle: catAxisRef.current?.value.trim() ?? '',
                            valAxisTitle: valAxisRef.current?.value.trim() ?? '',
                          })
                        }
                      >
                        {t('ribbonApply')}
                      </button>
                    </div>
                    <div className="rb-drop-title">{t('ribbonLegend')}</div>
                    <div className="rb-row">
                      {(
                        [
                          ['none', t('ribbonNone')],
                          ['b', t('ribbonLegendBottom')],
                          ['t', t('ribbonLegendTop')],
                          ['r', t('ribbonLegendRight')],
                          ['l', t('ribbonLegendLeft')],
                        ] as const
                      ).map(([pos, label]) => (
                        <button
                          key={pos}
                          className={`rb-icon ${(contextChartStyle?.legendPos ?? 'b') === pos ? 'active' : ''}`}
                          onClick={() => onEditChart?.({ legendPos: pos })}
                        >
                          {label}
                        </button>
                      ))}
                    </div>
                    <div className="rb-drop-title">{t('ribbonElementToggles')}</div>
                    <div className="rb-row">
                      <button
                        className={`rb-icon ${contextChartStyle?.dataLabels ? 'active' : ''}`}
                        onClick={() =>
                          onEditChart?.({ dataLabels: !contextChartStyle?.dataLabels })
                        }
                      >
                        {t('ribbonDataLabels')}
                      </button>
                      <button
                        className={`rb-icon ${contextChartStyle?.gridlines ? 'active' : ''}`}
                        onClick={() => onEditChart?.({ gridlines: !contextChartStyle?.gridlines })}
                      >
                        {t('ribbonGridlines')}
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupChartStyles')}>
              <div className="rb-drop-wrap">
                <button
                  className={`rb-big ${chartDrop === 'colors' ? 'active' : ''}`}
                  disabled={!onEditChart}
                  data-tip={t('ribbonChangeColorsTip')}
                  onMouseDown={(e) => {
                    e.stopPropagation()
                    closeSiblingPanels(e, closePanels, 'chart')
                  }}
                  onClick={() => setChartDrop((v) => (v === 'colors' ? null : 'colors'))}
                >
                  <span className="rb-big-icon" style={{ fontSize: 20 }}>
                    🎨
                    <RbCaret />
                  </span>
                  <span>{t('ribbonChangeColors')}</span>
                </button>
                {chartDrop === 'colors' && (
                  <div
                    className="rb-drop rb-chart-colors-drop"
                    onMouseDown={(e) => e.stopPropagation()}
                  >
                    {(
                      chartColorSchemes ??
                      CHART_COLOR_SCHEME_UI.map((s) => ({ ...s, label: t(s.label) }))
                    ).map((s) => (
                      <button
                        key={s.key}
                        className="rb-chart-scheme-card"
                        data-tip={s.label}
                        onClick={() => {
                          setChartDrop(null)
                          onEditChart?.({ colorScheme: s.key })
                        }}
                      >
                        <span className="rb-chart-scheme-swatches">
                          {s.colors.slice(0, 4).map((c) => (
                            <span
                              key={c}
                              className="rb-chart-scheme-dot"
                              style={{ background: c }}
                            />
                          ))}
                        </span>
                        <span className="rb-chart-scheme-label">{s.label}</span>
                      </button>
                    ))}
                  </div>
                )}
              </div>
              <div className="rb-chart-style-row">
                {CHART_STYLE_PRESETS.map((p) => (
                  <button
                    key={p.key}
                    className={`rb-chart-style-card ${chartPresetActive(contextChartStyle, p) ? 'active' : ''}`}
                    data-tip={t(p.label)}
                    disabled={!onEditChart}
                    onClick={() => onEditChart?.({ ...p.style })}
                  >
                    <ChartStyleThumb kind={contextChartStyle?.kind} style={p.style} />
                    <span className="rb-chart-scheme-label">{t(p.label)}</span>
                  </button>
                ))}
              </div>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupData')}>
              <button
                className="rb-big"
                data-tip={t('ribbonSwitchRowColTip')}
                disabled={!onEditChart}
                onClick={() => onEditChart?.({ switchRowCol: true })}
              >
                <span className="rb-big-icon">
                  <IconSwitchRowCol size={BIG} />
                </span>
                <span>{t('ribbonSwitchRowCol')}</span>
              </button>
              <button
                className="rb-big"
                data-tip={t('ribbonEditDataTip')}
                disabled={!onOpenChartDataDialog}
                onClick={onOpenChartDataDialog}
              >
                <span className="rb-big-icon">
                  <IconEditChartData size={BIG} />
                </span>
                <span>{t('ribbonEditData')}</span>
              </button>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupType')}>
              <button
                className="rb-big"
                disabled={!onEditChart}
                data-tip={t('ribbonChangeChartType')}
                onClick={() => setChartTypeDlgOpen(true)}
              >
                <span className="rb-big-icon">
                  <IconChangeChartType size={BIG} />
                </span>
                <span>{t('ribbonChangeChartType')}</span>
              </button>
            </Group>
            {chartTypeDlgOpen && (
              <ChartTypeDialog
                current={
                  contextChartStyle && contextChartStyle.kind !== 'unknown'
                    ? contextChartStyle.kind
                    : undefined
                }
                onClose={() => setChartTypeDlgOpen(false)}
                onConfirm={(kind) => {
                  setChartTypeDlgOpen(false)
                  onEditChart?.({ kind })
                }}
              />
            )}
          </>
        ) : tab === 'pictureFormat' ? (
          <>
            <Group label={t('ribbonGroupAdjust')}>
              <button
                className="rb-big"
                data-tip={
                  contextPictureCanCutout ? t('ribbonRemoveBgTip') : t('ribbonRemoveBgDisabledTip')
                }
                disabled={!onPictureCutout || !contextPictureCanCutout}
                onClick={onPictureCutout}
              >
                <span className="rb-big-icon">
                  {/* BIG, not BIG+2: the shared 28px glyph size keeps this button's
                      icon-row height (and label line) identical to its neighbors */}
                  <IconRemoveBg size={BIG} />
                </span>
                <span>{t('ribbonRemoveBg')}</span>
              </button>
              <div className="rb-drop-wrap">
                <button
                  className={`rb-big ${transparencyOpen ? 'active' : ''}`}
                  disabled={!onPictureOpacity || contextElementType !== 'picture'}
                  onMouseDown={(e) => {
                    e.stopPropagation()
                    closeSiblingPanels(e, closePanels, 'transparency')
                  }}
                  onClick={() => setTransparencyOpen((v) => !v)}
                  data-tip={t('ribbonTransparency')}
                >
                  <span className="rb-big-icon">
                    {/* 28px box around the 22px art so the icon row matches the SVG glyphs' height */}
                    <span className="ai-feature-icon" aria-hidden="true">
                      <img src={iconTransparency} width={22} height={22} alt="" />
                    </span>
                    <RbCaret />
                  </span>
                  <span>{t('ribbonTransparency')}</span>
                </button>
                {transparencyOpen && (
                  <div className="rb-drop rb-menu" onMouseDown={(e) => e.stopPropagation()}>
                    {[0, 15, 30, 50, 65, 80, 95].map((pct) => (
                      <button
                        key={pct}
                        onClick={() => {
                          setTransparencyOpen(false)
                          onPictureOpacity?.(1 - pct / 100)
                        }}
                      >
                        {pct}%
                      </button>
                    ))}
                  </div>
                )}
              </div>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('paneFormatOutline')}>
              <div className="rb-drop-wrap">
                <button
                  className={`rb-big ${pictureBorderOpen ? 'active' : ''}`}
                  disabled={!onPictureStroke}
                  onMouseDown={(e) => {
                    e.stopPropagation()
                    closeSiblingPanels(e, closePanels, 'pictureBorder')
                  }}
                  onClick={() => {
                    pictureBorderDraft.current = null
                    setPictureBorderOpen((v) => !v)
                  }}
                  data-tip={t('paneFormatOutline')}
                >
                  <span className="rb-big-icon">
                    <IconPageBorders size={BIG} />
                    <RbCaret />
                  </span>
                  <span>{t('paneFormatOutline')}</span>
                </button>
                {pictureBorderOpen && (
                  <div className="rb-drop rb-menu" onMouseDown={(e) => e.stopPropagation()}>
                    <label className="rb-menu-input">
                      {t('paneFormatOutlineColor')}
                      <input
                        type="color"
                        defaultValue={toPickerHex(contextPictureStroke?.color) ?? '#000000'}
                        onPointerDown={(e) => armColorInput(e.currentTarget)}
                        onChange={(e) => commitPictureBorder({ color: e.target.value })}
                      />
                    </label>
                    <div className="rb-menu-sep" />
                    {[0.5, 1, 1.5, 2.25, 3, 4.5, 6].map((pt) => (
                      <button
                        key={pt}
                        className={contextPictureStroke?.widthPt === pt ? 'active' : ''}
                        onClick={() => {
                          setPictureBorderOpen(false)
                          commitPictureBorder({ widthPt: pt }, true)
                        }}
                      >
                        {pt} pt
                      </button>
                    ))}
                    <div className="rb-menu-sep" />
                    <button
                      className={!contextPictureStroke ? 'active' : ''}
                      onClick={() => {
                        setPictureBorderOpen(false)
                        // A pending debounced color commit still holds the prior draft
                        // in its closure and would re-apply the border after the clear
                        if (pictureBorderTimer.current) {
                          window.clearTimeout(pictureBorderTimer.current)
                          pictureBorderTimer.current = null
                        }
                        pictureBorderDraft.current = null
                        onPictureStroke?.(null)
                      }}
                    >
                      {t('paneFormatNoOutline')}
                    </button>
                  </div>
                )}
              </div>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('ribbonGroupSize')}>
              <button
                className={`rb-big ${cropActive ? 'active' : ''}`}
                data-tip={t('ribbonCropTip')}
                disabled={!onPictureCrop || contextElementType !== 'picture'}
                onClick={onPictureCrop}
              >
                <span className="rb-big-icon">
                  <span className="ai-feature-icon" aria-hidden="true">
                    <img src={iconCrop} width={22} height={22} alt="" />
                  </span>
                </span>
                <span>{t('ribbonCrop')}</span>
              </button>
            </Group>
          </>
        ) : tab === 'shapeFormat' ? (
          <>
            <Group label={t('ribbonGroupShapeStyle')}>
              <div className="rb-drop-wrap">
                <button
                  className={`rb-big ${shapeStyleOpen ? 'active' : ''}`}
                  disabled={!onShapeStyle}
                  data-tip={t('ribbonShapeStyleTip')}
                  onMouseDown={(e) => {
                    e.stopPropagation()
                    closeSiblingPanels(e, closePanels, 'shapeStyle')
                  }}
                  onClick={() => setShapeStyleOpen((v) => !v)}
                >
                  <span className="rb-big-icon">
                    <IconShapeStyle size={BIG} />
                    <RbCaret />
                  </span>
                  <span>{t('ribbonGroupShapeStyle')}</span>
                </button>
                {shapeStyleOpen && (
                  <div
                    className="rb-drop rb-menu ctx-style-grid"
                    onMouseDown={(e) => e.stopPropagation()}
                  >
                    {RIBBON_SHAPE_STYLES.map((s, si) => (
                      <button
                        key={si}
                        className="ctx-style-cell"
                        style={{
                          background: s.fill,
                          borderColor: s.stroke,
                          borderStyle: s.dash ? 'dashed' : 'solid',
                        }}
                        aria-label={`${s.fill} / ${s.stroke}${s.dash ? ` (${s.dash})` : ''}`}
                        onClick={() => {
                          setShapeStyleOpen(false)
                          onShapeStyle?.(s)
                        }}
                      />
                    ))}
                  </div>
                )}
              </div>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('paneFormatFill')}>
              <div className="rb-drop-wrap">
                <button
                  className={`rb-big ${shapeFillOpen ? 'active' : ''}`}
                  disabled={!onShapeFill}
                  data-tip={t('paneFormatFill')}
                  onMouseDown={(e) => {
                    e.stopPropagation()
                    closeSiblingPanels(e, closePanels, 'shapeFill')
                  }}
                  onClick={() => setShapeFillOpen((v) => !v)}
                >
                  <span className="rb-big-icon">
                    <IconFillColor size={BIG} />
                    <RbCaret />
                  </span>
                  <span>{t('paneFormatFill')}</span>
                </button>
                {shapeFillOpen && (
                  <div className="rb-drop rb-menu" onMouseDown={(e) => e.stopPropagation()}>
                    <label className="rb-menu-input">
                      {t('paneFormatSolidFill')}
                      <input
                        type="color"
                        defaultValue="#ffffff"
                        onPointerDown={(e) => armColorInput(e.currentTarget)}
                        onChange={(e) => onShapeFill?.(e.target.value)}
                      />
                    </label>
                    <div className="rb-menu-sep" />
                    <div className="ctx-swatches-row">
                      {TEXT_COLORS.map((c) => (
                        <button
                          key={c}
                          className="ctx-swatch"
                          style={{ background: c }}
                          data-tip={c}
                          aria-label={c}
                          onClick={() => {
                            setShapeFillOpen(false)
                            onShapeFill?.(c)
                          }}
                        />
                      ))}
                    </div>
                    <div className="rb-menu-sep" />
                    <button
                      onClick={() => {
                        setShapeFillOpen(false)
                        onShapeFill?.('none')
                      }}
                    >
                      {t('paneFormatNoFill')}
                    </button>
                  </div>
                )}
              </div>
            </Group>
            <div className="ribbon-sep" />
            <Group label={t('paneFormatOutline')}>
              <div className="rb-drop-wrap">
                <button
                  className={`rb-big ${pictureBorderOpen ? 'active' : ''}`}
                  disabled={!onPictureStroke}
                  onMouseDown={(e) => {
                    e.stopPropagation()
                    closeSiblingPanels(e, closePanels, 'pictureBorder')
                  }}
                  onClick={() => {
                    pictureBorderDraft.current = null
                    setPictureBorderOpen((v) => !v)
                  }}
                  data-tip={t('paneFormatOutline')}
                >
                  <span className="rb-big-icon">
                    <IconPageBorders size={BIG} />
                    <RbCaret />
                  </span>
                  <span>{t('paneFormatOutline')}</span>
                </button>
                {pictureBorderOpen && (
                  <div className="rb-drop rb-menu" onMouseDown={(e) => e.stopPropagation()}>
                    <label className="rb-menu-input">
                      {t('paneFormatOutlineColor')}
                      <input
                        type="color"
                        defaultValue={toPickerHex(contextPictureStroke?.color) ?? '#000000'}
                        onPointerDown={(e) => armColorInput(e.currentTarget)}
                        onChange={(e) => commitPictureBorder({ color: e.target.value })}
                      />
                    </label>
                    <div className="rb-menu-sep" />
                    {[0.5, 1, 1.5, 2.25, 3, 4.5, 6].map((pt) => (
                      <button
                        key={pt}
                        className={contextPictureStroke?.widthPt === pt ? 'active' : ''}
                        onClick={() => {
                          setPictureBorderOpen(false)
                          commitPictureBorder({ widthPt: pt }, true)
                        }}
                      >
                        {pt} pt
                      </button>
                    ))}
                    <div className="rb-menu-sep" />
                    {(
                      [
                        ['solid', t('ribbonLineSolid')],
                        ['dash', t('ribbonLineDash')],
                        ['sysDot', t('ribbonLineDot')],
                        ['dashDot', t('ribbonLineDashDot')],
                      ] as const
                    ).map(([dash, label]) => (
                      <button
                        key={dash}
                        className={
                          contextPictureStroke &&
                          (dash === 'solid'
                            ? !contextPictureStroke.dashPreset
                            : contextPictureStroke.dashPreset === dash)
                            ? 'active'
                            : ''
                        }
                        onClick={() => {
                          setPictureBorderOpen(false)
                          onPictureStroke?.({
                            color: toPickerHex(contextPictureStroke?.color) ?? '#000000',
                            widthPt: contextPictureStroke?.widthPt ?? 1,
                            dash,
                          })
                        }}
                      >
                        {label}
                      </button>
                    ))}
                    <div className="rb-menu-sep" />
                    <button
                      className={!contextPictureStroke ? 'active' : ''}
                      onClick={() => {
                        setPictureBorderOpen(false)
                        // A pending debounced color commit still holds the prior draft
                        // in its closure and would re-apply the border after the clear
                        if (pictureBorderTimer.current) {
                          window.clearTimeout(pictureBorderTimer.current)
                          pictureBorderTimer.current = null
                        }
                        pictureBorderDraft.current = null
                        onPictureStroke?.(null)
                      }}
                    >
                      {t('paneFormatNoOutline')}
                    </button>
                  </div>
                )}
              </div>
            </Group>
          </>
        ) : null}
      </div>
    </div>
  )
}

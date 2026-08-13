/**
 * Preset data for the Insert tab: shape gallery / icon gallery / WordArt / SmartArt / chart samples / equation templates.
 * Display names go through i18n (lazy getters, follow language switches); prst/id/geometry data stay unchanged.
 */
import { SHAPE_GALLERY_GROUPS } from '@genoffice/ui'
import { t, type StringKey } from './i18n/locale'
import type { SmartArtLayout } from '../../../../packages/pptx-engine/src/smartart-layout'

// ── Shape gallery ─────────────────────────────────────────────────────────────

export interface ShapeDef {
  /** OOXML prstGeom name (also drives the gallery preview, see ShapePreview) */
  prst: string
  label: string
}

const shape = (prst: string, key: StringKey): ShapeDef => ({
  prst,
  get label() {
    return t(key)
  },
})

const shapeGroup = (key: StringKey, shapes: ShapeDef[]): { group: string; shapes: ShapeDef[] } => ({
  get group() {
    return t(key)
  },
  shapes,
})

/** Built from the cross-app shared gallery so every app shows the same shapes. */
export const SHAPE_GALLERY: Array<{ group: string; shapes: ShapeDef[] }> = SHAPE_GALLERY_GROUPS.map(
  (g) =>
    shapeGroup(
      g.groupKey as StringKey,
      g.shapes.map((sh) => shape(sh.prst, sh.labelKey as StringKey)),
    ),
)

// ── Icon gallery (24px stroke style, rasterized then inserted as PNG images) ──────────────

export interface IconDef {
  name: string
  /** Inner markup of viewBox 0 0 24 24 (stroke=currentColor fill=none) */
  body: string
}

const I = (key: StringKey, body: string): IconDef => ({
  get name() {
    return t(key)
  },
  body,
})

export const ICON_GALLERY: IconDef[] = [
  I(
    'ribbonIconStar',
    '<path d="M12 3l2.7 5.6 6.1.8-4.5 4.2 1.1 6L12 16.8 6.6 19.6l1.1-6L3.2 9.4l6.1-.8z"/>',
  ),
  I(
    'ribbonIconHeart',
    '<path d="M12 20.5C6 15.5 3 12.3 3 8.9 3 6.2 5.1 4 7.7 4c1.7 0 3.3.9 4.3 2.4C13 4.9 14.6 4 16.3 4 18.9 4 21 6.2 21 8.9c0 3.4-3 6.6-9 11.6z"/>',
  ),
  I('ribbonIconCheck', '<path d="M4 12.5l5 5L20 6.5"/>'),
  I('ribbonIconCross', '<path d="M6 6l12 12M18 6L6 18"/>'),
  I('ribbonIconPlus', '<path d="M12 5v14M5 12h14"/>'),
  I(
    'ribbonIconWarning',
    '<path d="M12 3L2.5 20h19z"/><path d="M12 9.5v5"/><circle cx="12" cy="17.2" r="0.6"/>',
  ),
  I(
    'ribbonIconInfo',
    '<circle cx="12" cy="12" r="9"/><path d="M12 11v5"/><circle cx="12" cy="8" r="0.6"/>',
  ),
  I(
    'ribbonIconQuestion',
    '<circle cx="12" cy="12" r="9"/><path d="M9.5 9.3A2.6 2.6 0 0 1 12 7.5c1.4 0 2.5 1 2.5 2.3 0 1.6-2.1 2-2.5 3.4v.6"/><circle cx="12" cy="16.6" r="0.6"/>',
  ),
  I(
    'ribbonIconBulb',
    '<path d="M9 18h6M10 21h4"/><path d="M12 3a6 6 0 0 1 3.5 10.9c-.8.6-1 1.3-1 2.1h-5c0-.8-.2-1.5-1-2.1A6 6 0 0 1 12 3z"/>',
  ),
  I(
    'ribbonIconGear',
    '<circle cx="12" cy="12" r="3.2"/><path d="M12 2.5v3M12 18.5v3M2.5 12h3M18.5 12h3M5 5l2.1 2.1M16.9 16.9L19 19M19 5l-2.1 2.1M7.1 16.9L5 19"/>',
  ),
  I('ribbonIconClock', '<circle cx="12" cy="12" r="9"/><path d="M12 6.5V12l3.8 2.4"/>'),
  I(
    'ribbonIconCalendar',
    '<rect x="3.5" y="5" width="17" height="15.5" rx="2"/><path d="M3.5 9.5h17M8 3v4M16 3v4"/>',
  ),
  I(
    'ribbonIconMail',
    '<rect x="3" y="5.5" width="18" height="13" rx="2"/><path d="M3.5 7l8.5 6 8.5-6"/>',
  ),
  I(
    'ribbonIconPhone',
    '<path d="M5.5 4h3.6l1.4 4.2-2.1 1.6a12.5 12.5 0 0 0 5.8 5.8l1.6-2.1L20 14.9v3.6c0 .8-.7 1.5-1.5 1.5C10.5 20 4 13.5 4 5.5 4 4.7 4.7 4 5.5 4z"/>',
  ),
  I(
    'ribbonIconLocation',
    '<path d="M12 21.5S5 14.7 5 9.8A7 7 0 0 1 19 9.8c0 4.9-7 11.7-7 11.7z"/><circle cx="12" cy="9.8" r="2.6"/>',
  ),
  I('ribbonIconUser', '<circle cx="12" cy="8" r="4"/><path d="M4.5 20.5a7.5 7.5 0 0 1 15 0"/>'),
  I(
    'ribbonIconTeam',
    '<circle cx="8.5" cy="9" r="3.2"/><circle cx="16.5" cy="10" r="2.6"/><path d="M2.5 19.5a6 6 0 0 1 12 0M14.8 19.5a5 5 0 0 1 6.7-4.7"/>',
  ),
  I(
    'ribbonIconHouse',
    '<path d="M4 11.5L12 4l8 7.5"/><path d="M6 10v10h12V10"/><path d="M10 20v-5.5h4V20"/>',
  ),
  I(
    'ribbonIconDocument',
    '<path d="M6.5 3h7L18.5 8v13h-12z"/><path d="M13 3.5V8.5h5"/><path d="M9 12.5h6M9 15.5h6"/>',
  ),
  I(
    'ribbonIconFolder',
    '<path d="M3.5 6.5c0-1.1.9-2 2-2h4l2 2.5h7c1.1 0 2 .9 2 2v9c0 1.1-.9 2-2 2h-13c-1.1 0-2-.9-2-2z"/>',
  ),
  I(
    'ribbonIconCamera',
    '<rect x="3" y="7" width="18" height="13" rx="2"/><path d="M8.5 7L10 4.5h4L15.5 7"/><circle cx="12" cy="13.3" r="3.4"/>',
  ),
  I(
    'ribbonIconMusic',
    '<path d="M9 18.5V6l10-2v12.5"/><circle cx="6.8" cy="18.5" r="2.3"/><circle cx="16.8" cy="16.5" r="2.3"/>',
  ),
  I(
    'ribbonIconCloud',
    '<path d="M7 18.5a4.2 4.2 0 0 1-.6-8.4 5.6 5.6 0 0 1 10.9 1.2A3.7 3.7 0 0 1 17 18.5z"/>',
  ),
  I(
    'ribbonIconLock',
    '<rect x="5.5" y="10.5" width="13" height="10" rx="2"/><path d="M8.5 10.5V8a3.5 3.5 0 0 1 7 0v2.5"/><circle cx="12" cy="15.3" r="1.2"/>',
  ),
  I('ribbonIconSearch', '<circle cx="10.5" cy="10.5" r="6.5"/><path d="M15.5 15.5L21 21"/>'),
  I(
    'ribbonIconCart',
    '<circle cx="9.5" cy="19.5" r="1.4"/><circle cx="17.5" cy="19.5" r="1.4"/><path d="M3 4.5h2.5l2.4 10.5h10.2L20.5 8h-13"/>',
  ),
  I(
    'ribbonIconTrophy',
    '<path d="M8 4h8v5.5a4 4 0 0 1-8 0z"/><path d="M8 5.5H4.5a3.5 3.5 0 0 0 3.6 3.5M16 5.5h3.5a3.5 3.5 0 0 1-3.6 3.5"/><path d="M12 13.5v3M8.5 20.5h7M10 16.5h4l.8 4h-5.6z"/>',
  ),
  I(
    'ribbonIconRocket',
    '<path d="M12 2.5c3.5 2 5 5.5 5 9l-2.5 5h-5L7 11.5c0-3.5 1.5-7 5-9z"/><circle cx="12" cy="9" r="1.8"/><path d="M9.5 16.5L7 21M14.5 16.5L17 21M12 17v4.5"/>',
  ),
]

/** IconDef → complete SVG string (rasterization input). */
export function iconSvg(def: IconDef, color: string, size = 512): string {
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 24 24" ` +
    `fill="none" stroke="${color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">` +
    `${def.body}</svg>`
  )
}

export const ICON_COLORS = ['#404040', '#4472C4', '#ED7D31', '#70AD47', '#C00000', '#7030A0']

// WordArt presets are shared across apps — see @genoffice/ui wordart-presets.ts.

// ── SmartArt layouts ──────────────────────────────────────────────────

export interface SmartArtDef {
  layout: SmartArtLayout
  label: string
  defaultItems: string[]
}

const smartArt = (
  layout: SmartArtDef['layout'],
  key: StringKey,
  items: StringKey[],
): SmartArtDef => ({
  layout,
  get label() {
    return t(key)
  },
  get defaultItems() {
    return items.map((k) => t(k))
  },
})

export const SMARTART_GALLERY: SmartArtDef[] = [
  smartArt('list', 'ribbonSmartArtList', [
    'ribbonSmartArtItem1',
    'ribbonSmartArtItem2',
    'ribbonSmartArtItem3',
  ]),
  smartArt('process', 'ribbonSmartArtProcess', [
    'ribbonSmartArtPlan',
    'ribbonSmartArtDo',
    'ribbonSmartArtCheck',
    'ribbonSmartArtAct',
  ]),
  smartArt('cycle', 'ribbonSmartArtCycle', [
    'ribbonSmartArtAnalyze',
    'ribbonSmartArtDesign',
    'ribbonSmartArtImplement',
    'ribbonSmartArtEvaluate',
  ]),
  smartArt('hierarchy', 'ribbonSmartArtHierarchy', [
    'ribbonSmartArtCeo',
    'ribbonSmartArtRnd',
    'ribbonSmartArtMarketing',
    'ribbonSmartArtOps',
  ]),
  smartArt('pyramid', 'ribbonSmartArtPyramid', [
    'ribbonSmartArtItem1',
    'ribbonSmartArtItem2',
    'ribbonSmartArtItem3',
  ]),
  smartArt('matrix', 'ribbonSmartArtMatrix', [
    'ribbonSmartArtStrength',
    'ribbonSmartArtWeakness',
    'ribbonSmartArtOpportunity',
    'ribbonSmartArtThreat',
  ]),
  smartArt('venn', 'ribbonSmartArtVenn', [
    'ribbonSmartArtItem1',
    'ribbonSmartArtItem2',
    'ribbonSmartArtItem3',
  ]),
]

// ── Chart sample data (default presets, editable after insertion) ────────────────

export interface ChartPresetDef {
  kind:
    | 'bar'
    | 'bar3D'
    | 'barStacked'
    | 'barPercentStacked'
    | 'barH'
    | 'line'
    | 'area'
    | 'pie'
    | 'pie3D'
    | 'doughnut'
    | 'scatter'
    | 'radar'
    | 'comboBarLine'
  label: string
}

const chart = (kind: ChartPresetDef['kind'], key: StringKey): ChartPresetDef => ({
  kind,
  get label() {
    return t(key)
  },
})

export const CHART_GALLERY: ChartPresetDef[] = [
  chart('bar', 'ribbonChartGalleryBar'),
  chart('barStacked', 'ribbonChartGalleryBarStacked'),
  chart('barPercentStacked', 'ribbonChartKindBarPercentStacked'),
  chart('barH', 'ribbonChartKindBarH'),
  chart('bar3D', 'ribbonChartKindBar3D'),
  chart('line', 'ribbonChartKindLine'),
  chart('area', 'ribbonChartKindArea'),
  chart('pie', 'ribbonChartKindPie'),
  chart('pie3D', 'ribbonChartKindPie3D'),
  chart('doughnut', 'ribbonChartGalleryDoughnut'),
  chart('scatter', 'ribbonChartKindScatter'),
  chart('radar', 'ribbonChartKindRadar'),
  chart('comboBarLine', 'ribbonChartKindCombo'),
]

export function chartSampleData(kind: ChartPresetDef['kind']): {
  categories: string[]
  series: Array<{ name: string; values: number[] }>
} {
  const quarters = ['ribbonSampleQ1', 'ribbonSampleQ2', 'ribbonSampleQ3', 'ribbonSampleQ4'] as const
  if (kind === 'pie' || kind === 'pie3D' || kind === 'doughnut') {
    return {
      categories: quarters.map((k) => t(k)),
      series: [{ name: t('ribbonSampleSales'), values: [8.2, 3.2, 1.4, 1.2] }],
    }
  }
  if (kind === 'scatter') {
    // Categories are the X values (numeric strings; the engine writes c:xVal)
    return {
      categories: ['0.7', '1.8', '2.6', '3.2', '4.5', '5.1'],
      series: [{ name: t('ribbonSampleYValues'), values: [2.7, 3.2, 0.8, 1.2, 1.9, 2.5] }],
    }
  }
  if (kind === 'comboBarLine') {
    // Combo chart: first two series as columns (sales/cost), last series as a line (growth rate, on the secondary axis)
    return {
      categories: quarters.map((k) => t(k)),
      series: [
        { name: t('ribbonSampleSales'), values: [4.3, 2.5, 3.5, 4.5] },
        { name: t('ribbonSampleCost'), values: [2.4, 1.8, 2.6, 2.8] },
        { name: t('ribbonSampleGrowth'), values: [2.0, 1.2, 2.6, 3.4] },
      ],
    }
  }
  if (kind === 'radar') {
    return {
      categories: [
        t('ribbonSampleComm'),
        t('ribbonSampleCollab'),
        t('ribbonSampleExecution'),
        t('ribbonSampleInnovation'),
        t('ribbonSampleLeadership'),
      ],
      series: [
        { name: t('ribbonSampleSeriesN', { n: 1 }), values: [3.5, 4.2, 3.0, 4.5, 3.8] },
        { name: t('ribbonSampleSeriesN', { n: 2 }), values: [4.4, 2.5, 4.0, 2.8, 3.2] },
      ],
    }
  }
  return {
    categories: [1, 2, 3, 4].map((n) => t('ribbonSampleCategoryN', { n })),
    series: [
      { name: t('ribbonSampleSeriesN', { n: 1 }), values: [4.3, 2.5, 3.5, 4.5] },
      { name: t('ribbonSampleSeriesN', { n: 2 }), values: [2.4, 4.4, 1.8, 2.8] },
      { name: t('ribbonSampleSeriesN', { n: 3 }), values: [2, 2, 3, 5] },
    ],
  }
}

// ── Equation templates (Unicode math text, inserted as Cambria Math text boxes) ─────────────

export interface EquationDef {
  label: string
  text: string
}

const eq = (key: StringKey, text: string): EquationDef => ({
  text,
  get label() {
    return t(key)
  },
})

export const EQUATION_GALLERY: EquationDef[] = [
  eq('ribbonEqQuadratic', 'x = (−b ± √(b² − 4ac)) / 2a'),
  eq('ribbonEqPythagorean', 'a² + b² = c²'),
  eq('ribbonEqEuler', 'e^(iπ) + 1 = 0'),
  eq('ribbonEqMassEnergy', 'E = mc²'),
  eq('ribbonEqCircleArea', 'A = πr²'),
  eq('ribbonEqArithSum', '∑ₖ₌₁ⁿ k = n(n+1)/2'),
  eq('ribbonEqGaussian', '∫₋∞^∞ e^(−x²) dx = √π'),
  eq('ribbonEqNaturalLog', 'lim(x→∞) (1 + 1/x)ˣ = e'),
  eq('ribbonEqMean', 'x̄ = (1/n) ∑ᵢ xᵢ'),
  eq('ribbonEqStdDev', 'σ = √( ∑(xᵢ − μ)² / n )'),
  eq('ribbonEqFourier', 'F(ω) = ∫₋∞^∞ f(t) e^(−iωt) dt'),
  eq('ribbonEqNormalDist', 'f(x) = (1/(σ√(2π))) e^(−(x−μ)²/2σ²)'),
]

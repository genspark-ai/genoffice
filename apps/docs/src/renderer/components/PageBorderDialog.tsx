import { useState, type CSSProperties } from 'react'
import type { SectionSettings } from '@genoffice/docx-engine'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'
import {
  pageBorderArtStripStyle,
  pageBorderLineCss,
  pageBorderStyleOf,
  type PageBorderSideName,
} from '../editor/pagination-gaps'

type Props = NonNullable<SectionSettings['pageBorderProps']>
type SideSpec = NonNullable<NonNullable<Props['sides']>['top']>
type Setting = 'none' | 'box' | 'shadow' | '3d' | 'custom'
export type PageBorderScope = 'whole' | 'section'

const SIDES: readonly PageBorderSideName[] = ['top', 'bottom', 'left', 'right']

/** Word's Style list order */
const LINE_STYLES = [
  'single',
  'dotted',
  'dashSmallGap',
  'dashed',
  'dotDash',
  'dotDotDash',
  'double',
  'triple',
  'thinThickSmallGap',
  'thickThinSmallGap',
  'thinThickThinSmallGap',
  'thinThickMediumGap',
  'thickThinMediumGap',
  'thinThickThinMediumGap',
  'thinThickLargeGap',
  'thickThinLargeGap',
  'thinThickThinLargeGap',
  'wave',
  'doubleWave',
  'dashDotStroked',
  'threeDEmboss',
  'threeDEngrave',
  'outset',
  'inset',
]

/** Word's Width list (pt) */
const LINE_WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6]

const ART_NAMES = [
  'apples',
  'archedScallops',
  'babyPacifier',
  'babyRattle',
  'balloons3Colors',
  'balloonsHotAir',
  'basicBlackDashes',
  'basicBlackDots',
  'basicBlackSquares',
  'basicThinLines',
  'basicWhiteDashes',
  'basicWhiteDots',
  'basicWhiteSquares',
  'basicWideInline',
  'basicWideMidline',
  'basicWideOutline',
  'bats',
  'birds',
  'birdsFlight',
  'cabins',
  'cakeSlice',
  'candyCorn',
  'celticKnotwork',
  'certificateBanner',
  'chainLink',
  'champagneBottle',
  'checkedBarBlack',
  'checkedBarColor',
  'checkered',
  'christmasTree',
  'circlesLines',
  'circlesRectangles',
  'classicalWave',
  'clocks',
  'compass',
  'confetti',
  'confettiGrays',
  'confettiOutline',
  'confettiStreamers',
  'confettiWhite',
  'cornerTriangles',
  'couponCutoutDashes',
  'couponCutoutDots',
  'crazyMaze',
  'creaturesButterfly',
  'creaturesFish',
  'creaturesInsects',
  'creaturesLadyBug',
  'crossStitch',
  'cup',
  'decoArch',
  'decoArchColor',
  'decoBlocks',
  'diamondsGray',
  'doubleD',
  'doubleDiamonds',
  'earth1',
  'earth2',
  'eclipsingSquares1',
  'eclipsingSquares2',
  'eggsBlack',
  'fans',
  'film',
  'firecrackers',
  'flowersBlockPrint',
  'flowersDaisies',
  'flowersModern1',
  'flowersModern2',
  'flowersPansy',
  'flowersRedRose',
  'flowersRoses',
  'flowersTeacup',
  'flowersTiny',
  'gems',
  'gingerbreadMan',
  'gradient',
  'handmade1',
  'handmade2',
  'heartBalloon',
  'heartGray',
  'hearts',
  'heebieJeebies',
  'holly',
  'houseFunky',
  'hypnotic',
  'iceCreamCones',
  'lightBulb',
  'lightning1',
  'lightning2',
  'mapPins',
  'mapleLeaf',
  'mapleMuffins',
  'marquee',
  'marqueeToothed',
  'moons',
  'mosaic',
  'musicNotes',
  'northwest',
  'ovals',
  'packages',
  'palmsBlack',
  'palmsColor',
  'paperClips',
  'papyrus',
  'partyFavor',
  'partyGlass',
  'pencils',
  'people',
  'peopleHats',
  'peopleWaving',
  'poinsettias',
  'postageStamp',
  'pumpkin1',
  'pushPinNote1',
  'pushPinNote2',
  'pyramids',
  'pyramidsAbove',
  'quadrants',
  'rings',
  'safari',
  'sawtooth',
  'sawtoothGray',
  'scaredCat',
  'seattle',
  'shadowedSquares',
  'shapes1',
  'shapes2',
  'sharksTeeth',
  'shorebirdTracks',
  'skyrocket',
  'snowflakeFancy',
  'snowflakes',
  'sombrero',
  'southwest',
  'stars',
  'stars3d',
  'starsBlack',
  'starsShadowed',
  'starsTop',
  'sun',
  'swirligig',
  'tornPaper',
  'tornPaperBlack',
  'trees',
  'triangle1',
  'triangle2',
  'triangleCircle1',
  'triangleCircle2',
  'triangleParty',
  'triangles',
  'tribal1',
  'tribal2',
  'tribal3',
  'tribal4',
  'tribal5',
  'tribal6',
  'twistedLines1',
  'twistedLines2',
  'vine',
  'waveline',
  'weavingAngles',
  'weavingBraid',
  'weavingRibbon',
  'weavingStrips',
  'whiteFlowers',
  'woodwork',
  'xIllusions',
  'zanyTriangles',
  'zigZag',
  'zigZagStitch',
]

const artLabel = (name: string): string =>
  name.replace(/([a-z])([A-Z0-9])/g, '$1 $2').replace(/^./, (c) => c.toUpperCase())

const SETTINGS: Array<[Setting, StringKey]> = [
  ['none', 'ribbonNone'],
  ['box', 'ribbonPbBox'],
  ['shadow', 'ribbonPbShadow'],
  ['3d', 'ribbonPb3D'],
  ['custom', 'ribbonPbCustom'],
]

const SIDE_KEYS: Record<PageBorderSideName, StringKey> = {
  top: 'ribbonMarginTop',
  bottom: 'ribbonMarginBottom',
  left: 'ribbonMarginLeft',
  right: 'ribbonMarginRight',
}

type Display = 'all' | 'firstPage' | 'notFirstPage'

interface Model {
  setting: Setting
  val: string
  art: string | null
  widthPt: number
  color: string | null
  /** per-side on/off (Custom); Box/Shadow/3-D keep all four */
  on: Record<PageBorderSideName, boolean>
  space: Record<PageBorderSideName, number>
  offsetFrom: 'page' | 'text'
  zOrderBack: boolean
  scope: PageBorderScope
  display: Display
}

const DEFAULT_SPACE = { top: 24, bottom: 24, left: 24, right: 24 }

function settingOf(sides: Props['sides']): Setting {
  const list = SIDES.map((s) => sides?.[s]).filter((s): s is SideSpec => !!s)
  if (list.length === 0) return 'none'
  const first = list[0]!
  const same = list.every(
    (s) =>
      s.val === first.val &&
      s.widthPt === first.widthPt &&
      s.color === first.color &&
      !!s.shadow === !!first.shadow &&
      !!s.frame === !!first.frame,
  )
  if (list.length < 4 || !same) return 'custom'
  if (first.shadow) return 'shadow'
  if (first.frame) return '3d'
  return 'box'
}

function modelOf(section: SectionSettings): Model {
  const p = section.pageBorder ? section.pageBorderProps : undefined
  const sides = p?.sides
  const first = SIDES.map((s) => sides?.[s]).find((s): s is SideSpec => !!s)
  const on = { top: false, bottom: false, left: false, right: false }
  const space = { ...DEFAULT_SPACE }
  for (const s of SIDES) {
    if (sides?.[s]) {
      on[s] = true
      space[s] = sides[s]!.spacePt
    }
  }
  const setting = section.pageBorder ? (p ? settingOf(sides) : 'box') : 'none'
  if (setting !== 'none' && setting !== 'custom') for (const s of SIDES) on[s] = true
  return {
    setting,
    val: first && !first.art ? first.val : 'single',
    art: first?.art ? first.val : null,
    widthPt: first?.widthPt ?? (first?.art ? 16 : 0.5),
    color: first?.color ?? null,
    on,
    space,
    // absent w:offsetFrom on a parsed border means text (pageBorderStyleOf)
    offsetFrom: p ? (p.offsetFrom ?? 'text') : 'page',
    zOrderBack: p?.zOrder === 'back',
    scope: p?.display ? 'section' : 'whole',
    display: p?.display ?? 'all',
  }
}

function sideOf(m: Model, side: PageBorderSideName): SideSpec {
  return {
    val: m.art ?? m.val,
    widthPt: m.widthPt,
    spacePt: m.space[side],
    ...(m.color ? { color: m.color } : {}),
    ...(m.art ? { art: true as const } : {}),
    ...(m.setting === 'shadow' ? { shadow: true as const } : {}),
    ...(m.setting === '3d' ? { frame: true as const } : {}),
  }
}

/** pageBorderProps for the model; null when the border is off */
export function pageBorderPropsOf(m: Model): Props | null {
  const active = SIDES.filter((s) => m.setting !== 'none' && m.on[s])
  if (active.length === 0) return null
  const sides: NonNullable<Props['sides']> = {}
  for (const s of active) sides[s] = sideOf(m, s)
  return {
    ...(m.display !== 'all' ? { display: m.display } : {}),
    offsetFrom: m.offsetFrom,
    ...(m.zOrderBack ? { zOrder: 'back' as const } : {}),
    spacePt: Math.max(...active.map((s) => m.space[s])),
    widthPt: m.widthPt,
    ...(m.color ? { color: m.color } : {}),
    sides,
  }
}

export function PageBorderDialog({
  section,
  onApply,
  onClose,
}: {
  section: SectionSettings
  onApply: (props: Props | null, scope: PageBorderScope) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [m, setM] = useState<Model>(() => modelOf(section))
  const [options, setOptions] = useState(false)
  const set = (patch: Partial<Model>) => setM((prev) => ({ ...prev, ...patch }))
  const allOn = { top: true, bottom: true, left: true, right: true }

  const pickSetting = (setting: Setting) =>
    set({ setting, ...(setting === 'none' || setting === 'custom' ? {} : { on: allOn }) })
  const toggleSide = (side: PageBorderSideName) => {
    const none = m.setting === 'none'
    const on = {
      ...(none ? { top: false, bottom: false, left: false, right: false } : m.on),
      [side]: none || !m.on[side],
    }
    const anyOn = SIDES.some((s) => on[s])
    set({ on, setting: anyOn ? 'custom' : 'none' })
  }
  const pickStyle = (val: string) => set({ val, art: null, ...(m.art ? { widthPt: 0.5 } : {}) })
  const pickArt = (art: string | null) =>
    set(art ? { art, widthPt: m.art ? m.widthPt : 16 } : { art: null, widthPt: 0.5 })

  const props = pageBorderPropsOf(m)
  const preview = pageBorderStyleOf({
    pageBorder: !!props,
    ...(props ? { pageBorderProps: props } : {}),
    marginTop: 1440,
    marginRight: 1440,
    marginBottom: 1440,
    marginLeft: 1440,
  })
  const inset = (side: PageBorderSideName) =>
    Math.min(28, Math.round((preview?.sides[side]?.insetPx ?? 0) / 3))
  const previewColor = `#${m.color ?? '000000'}`

  const submit = () => {
    onApply(props, m.scope)
    onClose()
  }

  const sideButton = (side: PageBorderSideName) => (
    <button
      key={side}
      type="button"
      className={`pb-side-btn pb-side-${side}${m.setting !== 'none' && m.on[side] ? ' active' : ''}`}
      aria-label={t(SIDE_KEYS[side])}
      aria-pressed={m.setting !== 'none' && m.on[side]}
      onClick={() => toggleSide(side)}
    >
      <span />
    </button>
  )

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form pb-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          submit()
        }}
      >
        <h2>{t('ribbonPbTitle')}</h2>
        <div className="pb-columns">
          <fieldset className="pb-col pb-settings">
            <legend>{t('ribbonPbSetting')}</legend>
            {SETTINGS.map(([key, label]) => (
              <button
                key={key}
                type="button"
                className={`pb-setting${m.setting === key ? ' active' : ''}`}
                aria-pressed={m.setting === key}
                onClick={() => pickSetting(key)}
              >
                <span className={`pb-setting-icon pb-setting-${key}`}>
                  <i />
                </span>
                {t(label)}
              </button>
            ))}
          </fieldset>
          <fieldset className="pb-col pb-line">
            <legend>{t('ribbonPbStyle')}</legend>
            <div className="pb-style-list" role="listbox" aria-label={t('ribbonPbStyle')}>
              {LINE_STYLES.map((val) => (
                <button
                  key={val}
                  type="button"
                  role="option"
                  aria-selected={!m.art && m.val === val}
                  title={artLabel(val)}
                  className={`pb-style${!m.art && m.val === val ? ' active' : ''}`}
                  onClick={() => pickStyle(val)}
                >
                  <span style={{ borderTop: pageBorderLineCss(val, 1.5, 'currentColor') }} />
                </button>
              ))}
            </div>
            <label className="fld">
              {t('ribbonPbColor')}
              <span className="pb-color-row">
                <input
                  type="color"
                  value={previewColor}
                  disabled={!!m.art}
                  aria-label={t('ribbonPbColor')}
                  onChange={(e) => set({ color: e.target.value.slice(1).toUpperCase() })}
                />
                <button
                  type="button"
                  className={`pb-auto${m.color ? '' : ' active'}`}
                  disabled={!!m.art}
                  onClick={() => set({ color: null })}
                >
                  {t('ribbonAutomatic')}
                </button>
              </span>
            </label>
            <label className="fld">
              {t('ribbonPbWidth')}
              {m.art ? (
                <input
                  type="number"
                  min={1}
                  max={31}
                  step={1}
                  value={m.widthPt}
                  onChange={(e) =>
                    set({ widthPt: Math.min(31, Math.max(1, Number(e.target.value) || 1)) })
                  }
                />
              ) : (
                <select
                  value={String(m.widthPt)}
                  onChange={(e) => set({ widthPt: Number(e.target.value) })}
                >
                  {(LINE_WIDTHS.includes(m.widthPt) ? LINE_WIDTHS : [...LINE_WIDTHS, m.widthPt])
                    .sort((a, b) => a - b)
                    .map((w) => (
                      <option key={w} value={String(w)}>
                        {`${w} pt`}
                      </option>
                    ))}
                </select>
              )}
            </label>
            <label className="fld">
              {t('ribbonPbArt')}
              <select value={m.art ?? ''} onChange={(e) => pickArt(e.target.value || null)}>
                <option value="">{t('ribbonPbArtNone')}</option>
                {ART_NAMES.map((name) => (
                  <option key={name} value={name}>
                    {artLabel(name)}
                  </option>
                ))}
              </select>
            </label>
          </fieldset>
          <fieldset className="pb-col pb-preview-col">
            <legend>{t('ribbonPbPreview')}</legend>
            <p className="modal-desc">{t('ribbonPbPreviewHint')}</p>
            <div className="pb-preview">
              <div className="pb-preview-side-col">
                {sideButton('top')}
                {sideButton('bottom')}
              </div>
              <div className="pb-preview-page-wrap">
                <div className="pb-preview-page">
                  {preview && (
                    <div
                      className="pb-preview-border"
                      style={{
                        top: inset('top'),
                        bottom: inset('bottom'),
                        left: inset('left'),
                        right: inset('right'),
                        borderTop: preview.sides.top?.css,
                        borderBottom: preview.sides.bottom?.css,
                        borderLeft: preview.sides.left?.css,
                        borderRight: preview.sides.right?.css,
                        boxShadow: preview.boxShadow,
                      }}
                    >
                      {SIDES.map((name) => {
                        const art = preview.sides[name]?.art
                        return art ? (
                          <div
                            key={name}
                            className="page-border-art"
                            style={pageBorderArtStripStyle(name, art) as CSSProperties}
                          />
                        ) : null
                      })}
                    </div>
                  )}
                  <div className="pb-preview-text" />
                </div>
                <div className="pb-preview-side-row">
                  {sideButton('left')}
                  {sideButton('right')}
                </div>
              </div>
            </div>
            <label className="fld">
              {t('ribbonPbApplyTo')}
              <select
                value={`${m.scope}:${m.display}`}
                onChange={(e) => {
                  const [scope, display] = e.target.value.split(':') as [PageBorderScope, Display]
                  set({ scope, display })
                }}
              >
                <option value="whole:all">{t('ribbonPbApplyWhole')}</option>
                <option value="section:all">{t('ribbonPbApplySection')}</option>
                <option value="section:firstPage">{t('ribbonPbApplyFirst')}</option>
                <option value="section:notFirstPage">{t('ribbonPbApplyNotFirst')}</option>
              </select>
            </label>
            <button
              type="button"
              className="pb-options-toggle"
              aria-expanded={options}
              onClick={() => setOptions((v) => !v)}
            >
              {t('ribbonPbOptions')}
            </button>
            {options && (
              <div className="pb-options">
                <div className="pb-margin-grid">
                  {SIDES.map((side) => (
                    <label key={side} className="fld">
                      {t(SIDE_KEYS[side])}
                      <input
                        type="number"
                        min={0}
                        max={31}
                        step={1}
                        value={m.space[side]}
                        onChange={(e) =>
                          set({
                            space: {
                              ...m.space,
                              [side]: Math.min(31, Math.max(0, Number(e.target.value) || 0)),
                            },
                          })
                        }
                      />
                    </label>
                  ))}
                </div>
                <label className="fld">
                  {t('ribbonPbMeasureFrom')}
                  <select
                    value={m.offsetFrom}
                    onChange={(e) => {
                      const offsetFrom = e.target.value as 'page' | 'text'
                      const d = offsetFrom === 'page' ? 24 : 4
                      set({ offsetFrom, space: { top: d, bottom: d, left: d, right: d } })
                    }}
                  >
                    <option value="text">{t('ribbonPbMeasureText')}</option>
                    <option value="page">{t('ribbonPbMeasurePage')}</option>
                  </select>
                </label>
                <label className="pb-check">
                  <input
                    type="checkbox"
                    checked={!m.zOrderBack}
                    onChange={(e) => set({ zOrderBack: !e.target.checked })}
                  />
                  {t('ribbonPbAlwaysFront')}
                </label>
              </div>
            )}
          </fieldset>
        </div>
        <div className="modal-actions">
          <button type="button" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button type="submit" className="btn-primary">
            {t('appOk')}
          </button>
        </div>
      </form>
    </div>
  )
}

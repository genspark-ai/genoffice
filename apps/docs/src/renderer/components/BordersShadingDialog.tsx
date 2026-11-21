import { useEffect, useRef, useState, type CSSProperties } from 'react'
import type { ParaBorderLine, SectionSettings } from '@genoffice/docx-engine'
import { NONE } from '../../shared/commands'
import {
  BORDER_SIDES,
  borderSpecOf,
  shadingDisplayOf,
  shadingSpecOf,
  type BorderSide,
  type BorderSpec,
  type ShadingSpec,
} from '../commands/borders-shading'
import { borderCssStyle } from '../editor/border-metrics'
import { paraBorderCss, paraBorderDropShadow } from '../editor/para-border-css'
import { SHADING_HATCHES, shadingHatchCss } from '../editor/shading-hatch'
import { parseRunBorderAttr } from '../editor/run-border'
import type { FontSubstitution } from '../font-check'
import { fontFamiliesFor, systemFamiliesBesidesCandidates } from '../font-list'
import { useI18n, type StringKey } from '../i18n/locale'
import { useSystemFontFamilies } from '../system-fonts'
import { useModalKeys } from './modal-keys'
import { PageBorderDialog, type PageBorderDraft, type PageBorderScope } from './PageBorderDialog'

type Tab = 'borders' | 'page' | 'shading'
const TABS: ReadonlyArray<{ id: Tab; label: StringKey }> = [
  { id: 'borders', label: 'ribbonParagraphBorders' },
  { id: 'page', label: 'ribbonBsPageBorder' },
  { id: 'shading', label: 'ribbonParagraphShading' },
]

type Setting = 'none' | 'box' | 'shadow' | '3d' | 'custom'
const SETTINGS: ReadonlyArray<{ id: Setting; label: StringKey }> = [
  { id: 'none', label: 'ribbonNone' },
  { id: 'box', label: 'ribbonPbBox' },
  { id: 'shadow', label: 'ribbonPbShadow' },
  { id: '3d', label: 'ribbonPb3D' },
  { id: 'custom', label: 'ribbonPbCustom' },
]
type BoxSetting = 'box' | 'shadow' | '3d'
const isBoxSetting = (s: Setting): s is BoxSetting => s === 'box' || s === 'shadow' || s === '3d'

const SIDE_LABELS: Record<BorderSide, StringKey> = {
  t: 'ribbonBorderTop',
  b: 'ribbonBorderBottom',
  l: 'ribbonBorderLeft',
  r: 'ribbonBorderRight',
}
const SIDE_NAMES = { t: 'top', b: 'bottom', l: 'left', r: 'right' } as const
const SPACE_LABELS: Record<BorderSide, StringKey> = {
  t: 'ribbonMarginTop',
  b: 'ribbonMarginBottom',
  l: 'ribbonMarginLeft',
  r: 'ribbonMarginRight',
}

/** the line styles the page draws as themselves; a double on text needs a width the mark does not give it */
const LINE_STYLES = ['single', 'dotted', 'dashed', 'double']
const TEXT_LINE_STYLES = ['single', 'dotted', 'dashed']

/** Word's Width list (pt) */
const LINE_WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6]
const DEFAULT_WIDTH = 0.5

/** the tints the page shows as the pattern colour blended over the fill */
const TINTS = [
  'pct5',
  'pct10',
  'pct12',
  'pct15',
  'pct20',
  'pct25',
  'pct30',
  'pct35',
  'pct37',
  'pct40',
  'pct45',
  'pct50',
  'pct55',
  'pct60',
  'pct62',
  'pct65',
  'pct70',
  'pct75',
  'pct80',
  'pct85',
  'pct87',
  'pct90',
  'pct95',
]
const HALF_TINTS = new Set(['pct12', 'pct37', 'pct62', 'pct87'])
const tintLabel = (val: string): string => `${val.slice(3)}${HALF_TINTS.has(val) ? '.5' : ''}%`

/** paragraph fills are document colours, the same in both themes */
const FILLS = [
  'FFFFFF',
  'F2F2F2',
  'D9D9D9',
  'BFBFBF',
  '7F7F7F',
  '000000',
  'FDE9D9',
  'FFF2CC',
  'E2EFDA',
  'DDEBF7',
  'E4DFEC',
  'FCE4D6',
  'F4B084',
  'FFD966',
  'A9D08E',
  '9BC2E6',
  'B4A7D6',
  'F8CBAD',
  'C00000',
  'FFC000',
  '00B050',
  '0070C0',
  '7030A0',
  'ED7D31',
]

const HATCH_LABELS: Record<(typeof SHADING_HATCHES)[number], StringKey> = {
  horzStripe: 'ribbonBsHatchHorz',
  vertStripe: 'ribbonBsHatchVert',
  reverseDiagStripe: 'ribbonBsHatchDownDiag',
  diagStripe: 'ribbonBsHatchUpDiag',
  horzCross: 'ribbonBsHatchGrid',
  diagCross: 'ribbonBsHatchTrellis',
  thinHorzStripe: 'ribbonBsHatchThinHorz',
  thinVertStripe: 'ribbonBsHatchThinVert',
  thinReverseDiagStripe: 'ribbonBsHatchThinDownDiag',
  thinDiagStripe: 'ribbonBsHatchThinUpDiag',
  thinHorzCross: 'ribbonBsHatchThinGrid',
  thinDiagCross: 'ribbonBsHatchThinTrellis',
}
const hatchLabel = (val: string): StringKey | undefined =>
  (HATCH_LABELS as Record<string, StringKey | undefined>)[val]

const nameLabel = (name: string): string =>
  name.replace(/([a-z])([A-Z0-9])/g, '$1 $2').replace(/^./, (c) => c.toUpperCase())

interface Pen {
  val: string
  color: string | null
  /** null: the line declares no width, which draws and saves as Word's half point */
  szPt: number | null
  /** Word's Shadow (w:shadow) or 3-D (w:frame) setting on the line */
  fx?: 'shadow' | '3d'
}

export interface BorderDraft {
  setting: Setting
  pen: Pen
  sides: Record<BorderSide, Pen | null>
  space: Record<BorderSide, number>
}

const DEFAULT_PEN: Pen = { val: 'single', color: null, szPt: null }
const samePen = (a: Pen, b: Pen): boolean =>
  a.val === b.val &&
  a.color === b.color &&
  (a.szPt ?? DEFAULT_WIDTH) === (b.szPt ?? DEFAULT_WIDTH) &&
  a.fx === b.fx

function settingOf(sides: BorderDraft['sides']): Setting {
  const drawn = BORDER_SIDES.map((s) => sides[s]).filter((p): p is Pen => !!p)
  if (drawn.length === 0) return 'none'
  if (drawn.length < 4 || !drawn.every((p) => samePen(p, drawn[0]!))) return 'custom'
  return drawn[0]!.fx ?? 'box'
}

/** the dialog's state for the borders a paragraph has */
export function borderDraftOf(spec: BorderSpec): BorderDraft {
  const sides = { t: null, b: null, l: null, r: null } as BorderDraft['sides']
  const space = { t: 0, b: 0, l: 0, r: 0 }
  for (const side of BORDER_SIDES) {
    const line = spec[side]
    if (!line) continue
    sides[side] = {
      val: line.val ?? 'single',
      color: line.color ?? null,
      szPt: line.szPt ?? null,
      ...(line.shadow ? { fx: 'shadow' as const } : line.frame ? { fx: '3d' as const } : {}),
    }
    space[side] = line.spacePt ?? 0
  }
  const first = BORDER_SIDES.map((s) => sides[s]).find((p) => !!p)
  return { setting: settingOf(sides), pen: first ?? DEFAULT_PEN, sides, space }
}

/** the borders the dialog's state draws */
export function borderSpecOfDraft(draft: BorderDraft): BorderSpec {
  const spec: BorderSpec = {}
  for (const side of BORDER_SIDES) {
    const pen = draft.sides[side]
    if (!pen) continue
    const line: ParaBorderLine = {}
    if (pen.color) line.color = pen.color
    if (pen.szPt) line.szPt = pen.szPt
    if (draft.space[side]) line.spacePt = draft.space[side]
    if (pen.val !== 'single') line.val = pen.val
    if (pen.fx === 'shadow') line.shadow = true
    if (pen.fx === '3d') line.frame = true
    spec[side] = line
  }
  return spec
}

/** a setting picked, the pen changed, or a side clicked, as Word's dialog takes each */
export function withSetting(draft: BorderDraft, setting: Setting): BorderDraft {
  if (setting === 'none')
    return { ...draft, setting, sides: { t: null, b: null, l: null, r: null } }
  if (isBoxSetting(setting)) {
    const { fx: _fx, ...plain } = draft.pen
    const pen: Pen = setting === 'box' ? plain : { ...plain, fx: setting }
    return { ...draft, setting, pen, sides: { t: pen, b: pen, l: pen, r: pen } }
  }
  return { ...draft, setting }
}

export function withPen(draft: BorderDraft, patch: Partial<Pen>): BorderDraft {
  const next = { ...draft, pen: { ...draft.pen, ...patch } }
  return isBoxSetting(draft.setting) ? withSetting(next, draft.setting) : next
}

export function withSideToggled(draft: BorderDraft, side: BorderSide): BorderDraft {
  const current = draft.sides[side]
  const sides = {
    ...draft.sides,
    [side]: current && samePen(current, draft.pen) ? null : draft.pen,
  }
  return { ...draft, sides, setting: BORDER_SIDES.some((s) => sides[s]) ? 'custom' : 'none' }
}

/** a character border as the four equal sides it is */
function textBorderSpecOf(bdr: unknown): BorderSpec {
  const parsed = typeof bdr === 'string' ? parseRunBorderAttr(bdr) : undefined
  if (!parsed || parsed.val === 'none' || parsed.val === 'nil') return {}
  const line: ParaBorderLine = { szPt: Math.max(parsed.sz, 2) / 8 }
  if (parsed.color && parsed.color !== 'auto') line.color = parsed.color
  if (parsed.val !== 'single') line.val = parsed.val
  return { t: line, b: line, l: line, r: line }
}

/** the section as the page border tab left it */
export function sectionWithPageBorder(
  section: SectionSettings,
  props: SectionSettings['pageBorderProps'] | null,
): SectionSettings {
  const { pageBorderProps: _drop, ...rest } = section
  return props
    ? { ...rest, pageBorder: true, pageBorderProps: props }
    : { ...rest, pageBorder: false }
}

export type BordersShadingCommand =
  'paragraph-borders' | 'paragraph-shading' | 'character-border' | 'character-shading'

type Target = 'paragraph' | 'text'

/**
 * Word's Format > Borders and Shading: the borders and the shading of the paragraphs
 * at the caret or of the selected text, with the page border between them as the tab
 * it is in Word.
 */
export function BordersShadingDialog({
  paragraph,
  text,
  styled,
  section,
  onCommand,
  onPageBorder,
  onHorizontalLine,
  onClose,
}: {
  /** the paragraph's `borders`, `borderLines`, `shadingFill` and `shadingPattern` attributes */
  paragraph: { borders: unknown; borderLines: unknown; fill: unknown; pattern: unknown }
  /** the selected text's `bdr` and `shading`; null when no text is selected */
  text: { bdr: unknown; shading: unknown } | null
  /** false where a paragraph carries only its sides and its fill (a text box, a header) */
  styled: boolean
  section: SectionSettings | null
  onCommand: (id: BordersShadingCommand, payload: string) => void
  onPageBorder: (props: SectionSettings['pageBorderProps'] | null, scope: PageBorderScope) => void
  /** left out where a line cannot go in (a text box, a header) */
  onHorizontalLine?: () => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [tab, setTab] = useState<Tab>('borders')
  const [target, setTarget] = useState<Target>('paragraph')
  const [opened] = useState(() => ({
    borders: {
      paragraph: borderDraftOf(borderSpecOf(paragraph.borders, paragraph.borderLines)),
      text: borderDraftOf(textBorderSpecOf(text?.bdr)),
    },
    shading: {
      paragraph: shadingSpecOf(paragraph.fill, paragraph.pattern),
      text: shadingSpecOf(text?.shading, null),
    },
  }))
  const [borders, setBorders] = useState(opened.borders)
  const [shadings, setShadings] = useState(opened.shading)
  const [options, setOptions] = useState(false)
  const draft = borders[target]
  const shading = shadings[target]
  const setDraft = (next: BorderDraft) => setBorders((prev) => ({ ...prev, [target]: next }))
  const setShading = (patch: Partial<ShadingSpec>) =>
    setShadings((prev) => ({ ...prev, [target]: { ...prev[target], ...patch } }))
  const onText = target === 'text'
  // a text box or header paragraph keeps only which sides are drawn and one fill
  const lineLive = onText || styled
  const patternLive = !onText && styled

  // the Page Border tab unmounts when another tab shows: what it held is kept here
  const pageDraftRef = useRef<PageBorderDraft | null>(null)
  const [pageDraft, setPageDraft] = useState<PageBorderDraft | null>(null)
  const showTab = (next: Tab) => {
    if (tab === 'page' && pageDraftRef.current?.touched) setPageDraft(pageDraftRef.current)
    setTab(next)
  }

  const tabRow = (
    <div className="bs-tabs" role="tablist">
      {TABS.map(({ id, label }) => (
        <button
          key={id}
          type="button"
          role="tab"
          aria-selected={tab === id}
          className={tab === id ? 'active' : ''}
          disabled={id === 'page' && !section}
          onClick={() => showTab(id)}
        >
          {t(label)}
        </button>
      ))}
    </div>
  )

  /** OK on any tab takes what the Borders and Shading tabs hold with it, as Word's one OK does */
  const applyTarget = () => {
    const spec = borderSpecOfDraft(draft)
    const drawn = BORDER_SIDES.filter((s) => spec[s])
    if (JSON.stringify(spec) !== JSON.stringify(borderSpecOfDraft(opened.borders[target]))) {
      if (onText) {
        const bdr = parseRunBorderAttr(String(text?.bdr ?? ''))
        onCommand(
          'character-border',
          drawn.length === 0
            ? NONE
            : JSON.stringify({
                val: draft.pen.val,
                sz: Math.round((draft.pen.szPt ?? DEFAULT_WIDTH) * 8),
                ...(draft.pen.color ? { color: draft.pen.color } : {}),
                ...(bdr?.space ? { space: bdr.space } : {}),
              }),
        )
      } else if (drawn.length === 0) onCommand('paragraph-borders', NONE)
      else onCommand('paragraph-borders', styled ? JSON.stringify(spec) : drawn.join(''))
    }
    if (JSON.stringify(shading) !== JSON.stringify(opened.shading[target])) {
      if (onText) onCommand('character-shading', shading.fill ?? NONE)
      else onCommand('paragraph-shading', styled ? JSON.stringify(shading) : (shading.fill ?? NONE))
    }
  }

  if (tab === 'page' && section)
    return (
      <PageBorderDialog
        section={pageDraft ? sectionWithPageBorder(section, pageDraft.props) : section}
        scope={pageDraft?.scope}
        draftRef={pageDraftRef}
        tabs={tabRow}
        onApply={(props, scope) => {
          applyTarget()
          onPageBorder(props, scope)
        }}
        onClose={onClose}
      />
    )

  const styles = onText ? TEXT_LINE_STYLES : LINE_STYLES
  const styleList = styles.includes(draft.pen.val) ? styles : [...styles, draft.pen.val]
  const width = draft.pen.szPt ?? DEFAULT_WIDTH
  const widths = LINE_WIDTHS.includes(width) ? LINE_WIDTHS : [...LINE_WIDTHS, width]
  const hatch = shading.pattern ? shadingHatchCss(shading.pattern, shading.color) : null
  const display = hatch ? shading.fill : shadingDisplayOf(shading)
  const previewStyle: CSSProperties = {
    backgroundColor: display ? `#${display}` : undefined,
    backgroundImage: hatch ?? undefined,
    boxShadow: paraBorderDropShadow(borderSpecOfDraft(draft)) ?? undefined,
  }
  for (const side of BORDER_SIDES) {
    const pen = draft.sides[side]
    if (!pen) continue
    const css = paraBorderCss({
      val: pen.val,
      szPt: pen.szPt ?? DEFAULT_WIDTH,
      ...(pen.color ? { color: pen.color } : {}),
      ...(pen.fx === '3d' ? { frame: true as const } : {}),
    })
    if (side === 't') previewStyle.borderTop = css
    else if (side === 'b') previewStyle.borderBottom = css
    else if (side === 'l') previewStyle.borderLeft = css
    else previewStyle.borderRight = css
  }
  const patterns = ['clear', 'solid', ...TINTS]
  const pattern = shading.pattern ?? 'clear'

  const sideButton = (side: BorderSide) => (
    <button
      key={side}
      type="button"
      data-side={side}
      className={`pb-side-btn pb-side-${SIDE_NAMES[side]}${draft.sides[side] ? ' active' : ''}`}
      aria-label={t(SIDE_LABELS[side])}
      aria-pressed={!!draft.sides[side]}
      disabled={onText}
      onClick={() => setDraft(withSideToggled(draft, side))}
    >
      <span />
    </button>
  )

  const preview = (
    <>
      <div className="pb-preview">
        <div className="pb-preview-side-col">
          {tab === 'borders' && sideButton('t')}
          {tab === 'borders' && sideButton('b')}
        </div>
        <div className="pb-preview-page-wrap">
          <div className="pb-preview-page bs-preview-page">
            <div className="bs-preview-para" style={previewStyle}>
              <span />
            </div>
          </div>
          {tab === 'borders' && (
            <div className="pb-preview-side-row">
              {sideButton('l')}
              {sideButton('r')}
            </div>
          )}
        </div>
      </div>
      <label className="fld">
        {t('ribbonPbApplyTo')}
        <select
          name="apply-to"
          value={target}
          onChange={(e) => setTarget(e.target.value as Target)}
        >
          <option value="paragraph">{t('ribbonSpParagraph')}</option>
          <option value="text" disabled={!text}>
            {t('ribbonPbMeasureText')}
          </option>
        </select>
      </label>
    </>
  )

  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form pb-dialog borders-shading-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          applyTarget()
          if (pageDraft) onPageBorder(pageDraft.props, pageDraft.scope)
        }}
      >
        {tabRow}
        {tab === 'borders' ? (
          <div className="pb-columns">
            <fieldset className="pb-col pb-settings">
              <legend>{t('ribbonPbSetting')}</legend>
              {SETTINGS.map(({ id, label }) => (
                <button
                  key={id}
                  type="button"
                  data-setting={id}
                  className={`pb-setting${draft.setting === id ? ' active' : ''}`}
                  aria-pressed={draft.setting === id}
                  // a character border and a text box paragraph carry plain lines only
                  disabled={
                    (onText && id === 'custom') ||
                    ((onText || !styled) && (id === 'shadow' || id === '3d'))
                  }
                  onClick={() => setDraft(withSetting(draft, id))}
                >
                  <span className={`pb-setting-icon pb-setting-${id}`}>
                    <i />
                  </span>
                  {t(label)}
                </button>
              ))}
            </fieldset>
            <fieldset className="pb-col pb-line" disabled={!lineLive}>
              <legend>{t('ribbonPbStyle')}</legend>
              <div className="pb-style-list" role="listbox" aria-label={t('ribbonPbStyle')}>
                {styleList.map((val) => (
                  <button
                    key={val}
                    type="button"
                    role="option"
                    data-line-style={val}
                    aria-selected={draft.pen.val === val}
                    title={nameLabel(val)}
                    className={`pb-style${draft.pen.val === val ? ' active' : ''}`}
                    onClick={() => setDraft(withPen(draft, { val }))}
                  >
                    {styles.includes(val) ? (
                      <span
                        style={{
                          borderTop: `${val === 'double' ? 3 : 2}px ${borderCssStyle(val)} currentColor`,
                        }}
                      />
                    ) : (
                      nameLabel(val)
                    )}
                  </button>
                ))}
              </div>
              <label className="fld">
                {t('ribbonPbColor')}
                <span className="pb-color-row">
                  <input
                    type="color"
                    name="border-color"
                    value={`#${draft.pen.color ?? '000000'}`}
                    aria-label={t('ribbonPbColor')}
                    onChange={(e) =>
                      setDraft(withPen(draft, { color: e.target.value.slice(1).toUpperCase() }))
                    }
                  />
                  <button
                    type="button"
                    className={`pb-auto${draft.pen.color ? '' : ' active'}`}
                    onClick={() => setDraft(withPen(draft, { color: null }))}
                  >
                    {t('ribbonAutomatic')}
                  </button>
                </span>
              </label>
              <label className="fld">
                {t('ribbonPbWidth')}
                <select
                  name="border-width"
                  value={String(width)}
                  onChange={(e) => setDraft(withPen(draft, { szPt: Number(e.target.value) }))}
                >
                  {[...widths]
                    .sort((a, b) => a - b)
                    .map((w) => (
                      <option key={w} value={String(w)}>
                        {`${w} pt`}
                      </option>
                    ))}
                </select>
              </label>
            </fieldset>
            <fieldset className="pb-col pb-preview-col">
              <legend>{t('ribbonPbPreview')}</legend>
              <p className="modal-desc">{t('ribbonPbPreviewHint')}</p>
              {preview}
              <button
                type="button"
                className="pb-options-toggle"
                aria-expanded={options}
                disabled={onText || !styled}
                onClick={() => setOptions((v) => !v)}
              >
                {t('ribbonPbOptions')}
              </button>
              {options && !onText && styled && (
                <div className="pb-options">
                  <p className="modal-desc">{t('ribbonLnFromText')}</p>
                  <div className="pb-margin-grid">
                    {BORDER_SIDES.map((side) => (
                      <label key={side} className="fld">
                        {t(SPACE_LABELS[side])}
                        <input
                          type="number"
                          name={`space-${side}`}
                          min={0}
                          max={31}
                          step={1}
                          value={draft.space[side]}
                          onChange={(e) =>
                            setDraft({
                              ...draft,
                              space: {
                                ...draft.space,
                                [side]: Math.min(
                                  31,
                                  Math.max(0, Math.round(Number(e.target.value) || 0)),
                                ),
                              },
                            })
                          }
                        />
                      </label>
                    ))}
                  </div>
                </div>
              )}
            </fieldset>
          </div>
        ) : (
          <div className="pb-columns bs-shading-columns">
            <fieldset className="pb-col">
              <legend>{t('ribbonBsFill')}</legend>
              <div className="bs-fills">
                {FILLS.map((color) => (
                  <button
                    key={color}
                    type="button"
                    data-fill={color}
                    aria-label={`#${color}`}
                    className={shading.fill?.toUpperCase() === color ? 'active' : ''}
                    style={{ background: `#${color}` }}
                    onClick={() => setShading({ fill: color })}
                  />
                ))}
              </div>
              <span className="pb-color-row">
                <input
                  type="color"
                  name="fill-color"
                  value={`#${shading.fill ?? 'FFFFFF'}`}
                  aria-label={t('ribbonBsFill')}
                  onChange={(e) => setShading({ fill: e.target.value.slice(1).toUpperCase() })}
                />
                <button
                  type="button"
                  data-fill="none"
                  className={`pb-auto${shading.fill ? '' : ' active'}`}
                  onClick={() => setShading({ fill: null })}
                >
                  {t('ribbonNoColor')}
                </button>
              </span>
            </fieldset>
            <fieldset className="pb-col" disabled={!patternLive}>
              <legend>{t('ribbonBsPatterns')}</legend>
              <label className="fld">
                {t('ribbonPbStyle')}
                <select
                  name="pattern"
                  value={pattern}
                  onChange={(e) =>
                    setShading({ pattern: e.target.value === 'clear' ? null : e.target.value })
                  }
                >
                  {(patterns.includes(pattern) || hatchLabel(pattern)
                    ? patterns
                    : [...patterns, pattern]
                  ).map((val) => (
                    <option key={val} value={val}>
                      {val === 'clear'
                        ? t('ribbonBsPatternClear')
                        : val === 'solid'
                          ? t('ribbonBsPatternSolid')
                          : val.startsWith('pct')
                            ? tintLabel(val)
                            : nameLabel(val)}
                    </option>
                  ))}
                  <optgroup label={t('ribbonBsHatches')}>
                    {SHADING_HATCHES.map((val) => (
                      <option key={val} value={val}>
                        {t(HATCH_LABELS[val])}
                      </option>
                    ))}
                  </optgroup>
                </select>
              </label>
              <label className="fld">
                {t('ribbonPbColor')}
                <span className="pb-color-row">
                  <input
                    type="color"
                    name="pattern-color"
                    value={`#${shading.color ?? '000000'}`}
                    disabled={!shading.pattern}
                    aria-label={t('ribbonPbColor')}
                    onChange={(e) => setShading({ color: e.target.value.slice(1).toUpperCase() })}
                  />
                  <button
                    type="button"
                    className={`pb-auto${shading.color ? '' : ' active'}`}
                    disabled={!shading.pattern}
                    onClick={() => setShading({ color: null })}
                  >
                    {t('ribbonAutomatic')}
                  </button>
                </span>
              </label>
            </fieldset>
            <fieldset className="pb-col pb-preview-col">
              <legend>{t('ribbonPbPreview')}</legend>
              {preview}
            </fieldset>
          </div>
        )}
        <div className="modal-actions">
          <button
            type="button"
            name="horizontalLine"
            className="bs-horizontal-line"
            disabled={!onHorizontalLine}
            onClick={() => {
              onClose()
              onHorizontalLine?.()
            }}
          >
            {t('ribbonBsHorizontalLine')}
          </button>
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

/** Word's Format > Font Substitution: the document's fonts this computer lacks, and what stands in */
export function FontSubstitutionDialog({
  fonts,
  onSubstitute,
  onConvert,
  onClose,
}: {
  fonts: FontSubstitution[]
  /** resolves false when the picked font cannot stand in; null hands the choice back to the app */
  onSubstitute: (name: string, substitute: string | null) => Promise<boolean>
  /** Convert Permanently: missing font -> the font the document names from now on */
  onConvert: (fonts: Record<string, string>) => void
  onClose: () => void
}) {
  const { t, lang } = useI18n()
  const keys = useModalKeys(onClose)
  const { families: systemFamilies, load: loadSystemFonts } = useSystemFontFamilies()
  useEffect(() => loadSystemFonts(), [loadSystemFonts])
  const candidates = fontFamiliesFor(lang)
  const families = [...candidates, ...systemFamiliesBesidesCandidates(candidates, systemFamilies)]
  // the picks made here, over what the dialog opened with: '' is the app's own choice
  const [picks, setPicks] = useState<Record<string, string>>({})
  const pickOf = (font: FontSubstitution) =>
    picks[font.name] ?? (font.chosen ? (font.substitute ?? '') : '')
  const standIn = (font: FontSubstitution) => pickOf(font) || (font.chosen ? null : font.substitute)
  const convertible = Object.fromEntries(
    fonts.flatMap((font) => {
      const to = standIn(font)
      return to ? [[font.name, to]] : []
    }),
  )
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal gs-form tco-dialog doc-dialog font-substitution-dialog">
        <h2>{t('ribbonFontSubTitle')}</h2>
        {fonts.length === 0 ? (
          <p className="font-substitution-none">{t('ribbonFontSubNone')}</p>
        ) : (
          <div className="recipients-table">
            <table>
              <thead>
                <tr>
                  <th>{t('ribbonFontSubMissing')}</th>
                  <th>{t('ribbonFontSubSubstitute')}</th>
                </tr>
              </thead>
              <tbody>
                {fonts.map((font) => {
                  const pick = pickOf(font)
                  const auto = font.chosen ? null : font.substitute
                  return (
                    <tr key={font.name} data-font={font.name}>
                      <td>{font.name}</td>
                      <td>
                        <select
                          value={pick}
                          aria-label={`${t('ribbonFontSubSubstitute')}: ${font.name}`}
                          onChange={(e) => {
                            const value = e.target.value
                            void onSubstitute(font.name, value || null).then((ok) => {
                              if (ok) setPicks((now) => ({ ...now, [font.name]: value }))
                            })
                          }}
                        >
                          <option value="">
                            {auto
                              ? `${t('ribbonFontSubDefault')} (${auto})`
                              : t('ribbonFontSubDefault')}
                          </option>
                          {pick && !families.includes(pick) && <option value={pick}>{pick}</option>}
                          {families.map((family) => (
                            <option key={family} value={family}>
                              {family}
                            </option>
                          ))}
                        </select>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        )}
        <div className="modal-actions">
          <button
            type="button"
            name="convert"
            disabled={Object.keys(convertible).length === 0}
            onClick={() => {
              onClose()
              onConvert(convertible)
            }}
          >
            {t('ribbonFontSubConvert')}
          </button>
          <div className="tco-actions-right">
            <button type="button" className="primary" onClick={onClose}>
              {t('ribbonClose')}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}

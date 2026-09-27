import { useEffect, useState } from 'react'
import type { Editor } from '@tiptap/core'
import type { Selection, Transaction } from '@tiptap/pm/state'
import { Dropdown, isSymbolFontFamily, type DropdownOption } from '@genoffice/ui'
import { useI18n, type StringKey } from '../i18n/locale'
import { fontFamiliesFor } from '../font-list'
import { fontSizeLabel, fontSizeOptions, parseFontSize } from '../font-sizes'
import { useSystemFontFamilies } from '../system-fonts'
import { charScaleEm, cssFontFamily, wordKerns } from '../line-metrics'
import { useModalKeys } from './modal-keys'

/**
 * Word's Font dialog: the Font tab (faces, style, size, color, effects) and the
 * Advanced tab (character scale / spacing / position / kerning). Fields the user
 * never touches are left out of the applied attrs, so a mixed selection keeps
 * its per-run values.
 */

const SCALE_PRESETS = [33, 50, 66, 80, 90, 100, 150, 200]

const FONT_STYLES: Array<{ key: string; nameKey: StringKey }> = [
  { key: 'regular', nameKey: 'appFontRegular' },
  { key: 'italic', nameKey: 'appFontItalic' },
  { key: 'bold', nameKey: 'appFontBold' },
  { key: 'boldItalic', nameKey: 'appFontBoldItalic' },
]

type SpacingMode = 'normal' | 'expanded' | 'condensed'
type PositionMode = 'normal' | 'raised' | 'lowered'
type Tab = 'font' | 'advanced'

/** Advanced-tab state derived from a run's docTextStyle attrs (Word's dialog fields) */
export interface CharSpacingState {
  scalePct: number
  spacing: SpacingMode
  spacingPt: number
  position: PositionMode
  positionPt: number
  kern: boolean
  kernPt: number
}

export function charSpacingFromAttrs(
  attrs: Record<string, unknown>,
  fallbackSizePt: number,
): CharSpacingState {
  const twips = attrs.charSpacingTwips
  const spacingTwips = typeof twips === 'number' ? twips : 0
  const half = attrs.positionHalfPoints
  const positionHalf = typeof half === 'number' ? half : 0
  const kernHalf = attrs.kernHalfPoints
  const scale = attrs.charScalePct
  return {
    scalePct: typeof scale === 'number' && scale > 0 ? scale : 100,
    spacing: spacingTwips > 0 ? 'expanded' : spacingTwips < 0 ? 'condensed' : 'normal',
    spacingPt: spacingTwips ? Math.abs(spacingTwips) / 20 : 1,
    position: positionHalf > 0 ? 'raised' : positionHalf < 0 ? 'lowered' : 'normal',
    positionPt: positionHalf ? Math.abs(positionHalf) / 2 : 3,
    kern: typeof kernHalf === 'number' && kernHalf > 0,
    kernPt: typeof kernHalf === 'number' && kernHalf > 0 ? kernHalf / 2 : fallbackSizePt,
  }
}

/** docTextStyle attrs for the advanced fields the user touched */
export function charSpacingAttrs(
  state: CharSpacingState,
  touched: ReadonlySet<string>,
  prior: Record<string, unknown>,
  sizeHalfPoints: number,
): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  if (touched.has('scale')) {
    const pct = Math.round(state.scalePct)
    out.charScalePct = pct > 0 && pct !== 100 ? pct : null
    // the display twins are recomputed per text node by setTextStyleWithScale
    out.charScaleX = null
    out.charScaleEm = null
  }
  if (touched.has('spacing')) {
    const twips = Math.round(state.spacingPt * 20)
    out.charSpacingTwips =
      state.spacing === 'normal' || !twips ? null : state.spacing === 'condensed' ? -twips : twips
  }
  if (touched.has('position')) {
    const half = Math.round(state.positionPt * 2)
    out.positionHalfPoints =
      state.position === 'normal' || !half ? null : state.position === 'lowered' ? -half : half
  }
  if (touched.has('kern')) {
    const half = state.kern ? Math.max(1, Math.round(state.kernPt * 2)) : null
    // an explicit 0 switches kerning off when the style asked for it
    const kernHalfPoints = half ?? (prior.kernHalfPoints != null ? 0 : null)
    out.kernHalfPoints = kernHalfPoints
    out.kern = wordKerns(kernHalfPoints ?? undefined, sizeHalfPoints) ?? null
  }
  return out
}

/**
 * setMark('docTextStyle') for the dialog: merges `attrs` over each text node's
 * existing mark in every selection range (a table CellSelection is several) and,
 * when the scale changed, recomputes that node's w:w letter-spacing
 * approximation in the same step, so one undo reverts it all.
 */
export function setTextStyleWithScale(
  tr: Transaction,
  ranges: ReadonlyArray<{ from: number; to: number }>,
  attrs: Record<string, unknown>,
): boolean {
  const type = tr.doc.type.schema.marks.docTextStyle
  if (ranges.every((r) => r.from === r.to)) {
    // same base as setMark: the pending stored marks, else the marks at the caret
    const base = (tr.storedMarks ?? tr.selection.$from.marks()).find((m) => m.type === type)
    const merged = { ...(base?.attrs ?? {}), ...attrs }
    if ('charScalePct' in attrs) {
      // text not typed yet: assume Latin glyph widths
      const pct = merged.charScalePct as number | null
      merged.charScaleEm = pct && pct !== 100 ? charScaleEm('x', pct) : null
    }
    tr.addStoredMark(type.create(merged))
    return true
  }
  for (const { from, to } of ranges) {
    tr.doc.nodesBetween(from, to, (node, pos) => {
      if (!node.isText || !node.text) return
      const merged = { ...(node.marks.find((m) => m.type === type)?.attrs ?? {}), ...attrs }
      if ('charScalePct' in attrs) {
        const pct = merged.charScalePct as number | null
        merged.charScaleEm = pct && pct !== 100 ? charScaleEm(node.text, pct) : null
      }
      tr.addMark(Math.max(pos, from), Math.min(pos + node.nodeSize, to), type.create(merged))
    })
  }
  return true
}

/** the selection's ranges as plain positions (CellSelection = one per cell) */
export function selectionRanges(sel: Selection): Array<{ from: number; to: number }> {
  return sel.ranges.map((r) => ({ from: r.$from.pos, to: r.$to.pos }))
}

/** Font-tab values the dialog edits; `touched` keys: fontLatin, fontEastAsia, size, color, vertAlign */
export interface FontTabState {
  fontLatin: string
  fontEastAsia: string
  sizePt: number
  colorHex: string
  vertAlign: string
}

/** docTextStyle attrs for the Font-tab fields the user touched */
export function fontTabAttrs(
  state: FontTabState,
  touched: ReadonlySet<string>,
): Record<string, unknown> {
  const out: Record<string, unknown> = {}
  // each picker writes only its own rFonts slot; empty means inherit
  if (touched.has('fontLatin')) out.fontAscii = state.fontLatin || null
  if (touched.has('fontEastAsia')) {
    out.font = state.fontEastAsia || null
    out.eastAsiaFont = state.fontEastAsia || null
    out.eaSlotEmpty = state.fontEastAsia ? false : null
  }
  if (touched.has('size')) out.sizeHalfPoints = Math.round(state.sizePt * 2)
  if (touched.has('color')) out.color = state.colorHex.replace('#', '').toUpperCase()
  if (touched.has('vertAlign')) out.vertAlign = state.vertAlign || null
  return out
}

export function FontDialog({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t, lang } = useI18n()
  const modalKeys = useModalKeys(onClose)
  const fontFamilies = fontFamiliesFor(lang)
  const sizeOptions = fontSizeOptions(lang)
  const { families: systemFontFamilies, load: loadSystemFonts } = useSystemFontFamilies()
  // the dialog opens from a click, so activation is still live here
  useEffect(() => loadSystemFonts(), [loadSystemFonts])
  const textAttrs = editor.getAttributes('docTextStyle')
  const initialStyle = editor.isActive('bold')
    ? editor.isActive('italic')
      ? 'boldItalic'
      : 'bold'
    : editor.isActive('italic')
      ? 'italic'
      : 'regular'

  const [tab, setTab] = useState<Tab>('font')
  const [fontEastAsia, setFontEastAsia] = useState((textAttrs.font as string | null) ?? '')
  const [fontLatin, setFontLatin] = useState((textAttrs.fontAscii as string | null) ?? '')
  const [size, setSize] = useState(
    textAttrs.sizeHalfPoints ? Number(textAttrs.sizeHalfPoints) / 2 : 11,
  )
  const [style, setStyle] = useState<string>(initialStyle)
  const [color, setColor] = useState(`#${(textAttrs.color as string | null) ?? '000000'}`)
  const [underline, setUnderline] = useState(editor.isActive('underline'))
  const [strike, setStrike] = useState(editor.isActive('strike'))
  const [vertAlign, setVertAlign] = useState<string>((textAttrs.vertAlign as string | null) ?? '')
  const [dstrike, setDstrike] = useState(textAttrs.dstrike === true)
  const [caps, setCaps] = useState<'all' | 'small' | ''>(
    textAttrs.caps === 'all' || textAttrs.caps === 'small' ? textAttrs.caps : '',
  )
  const [hidden, setHidden] = useState(textAttrs.vanish === true)
  const [adv, setAdv] = useState<CharSpacingState>(() => charSpacingFromAttrs(textAttrs, size))
  // only touched fields are written: the initial values are one run's snapshot,
  // and a mixed selection must keep every other run's own value
  const [touched, setTouched] = useState<ReadonlySet<string>>(() => new Set())
  const touch = (key: string) => setTouched((prev) => new Set(prev).add(key))
  const field =
    <T,>(key: string, set: (v: T) => void) =>
    (v: T) => {
      set(v)
      touch(key)
    }
  const setAdvField = <K extends keyof CharSpacingState>(
    group: string,
    key: K,
    value: CharSpacingState[K],
  ) => {
    setAdv((prev) => ({ ...prev, [key]: value }))
    touch(group)
  }

  const apply = () => {
    if (!editor.isEditable) {
      onClose()
      return
    }
    const sizeHalfPoints = Math.round(size * 2)
    const effects: Record<string, unknown> = {}
    if (touched.has('dstrike')) effects.dstrike = dstrike ? true : textAttrs.dstrike ? false : null
    if (touched.has('caps')) effects.caps = caps || (textAttrs.caps ? 'none' : null)
    if (touched.has('hidden')) {
      effects.vanish = hidden ? true : textAttrs.vanish ? false : null
      effects.vanishOwn = hidden ? true : textAttrs.vanish ? false : null
    }
    const styleAttrs: Record<string, unknown> = {
      ...fontTabAttrs(
        { fontLatin, fontEastAsia, sizePt: size, colorHex: color, vertAlign },
        touched,
      ),
      ...effects,
      ...charSpacingAttrs(adv, touched, textAttrs, sizeHalfPoints),
    }
    const ranges = selectionRanges(editor.state.selection)
    let chain = editor.chain().focus()
    if (Object.keys(styleAttrs).length > 0)
      chain = chain.command(({ tr }) => setTextStyleWithScale(tr, ranges, styleAttrs))
    if (touched.has('style')) {
      const wantBold = style === 'bold' || style === 'boldItalic'
      const wantItalic = style === 'italic' || style === 'boldItalic'
      chain = wantBold ? chain.setMark('bold') : chain.unsetMark('bold')
      chain = wantItalic ? chain.setMark('italic') : chain.unsetMark('italic')
    }
    if (touched.has('underline'))
      chain = underline ? chain.setMark('underline') : chain.unsetMark('underline')
    if (touched.has('strike')) chain = strike ? chain.setMark('strike') : chain.unsetMark('strike')
    chain.run()
    onClose()
  }

  const fontPicker = (value: string, onPick: (v: string) => void, label: string) => (
    <label>
      {label}
      <Dropdown
        value={value}
        ariaLabel={label}
        options={[
          ...fontFamilies.map((f): DropdownOption => ({
            value: f,
            label: f,
            render: <span style={{ fontFamily: cssFontFamily(f) }}>{f}</span>,
          })),
          ...systemFontFamilies.map((f): DropdownOption => ({
            value: f,
            label: f,
            render: (
              // symbol fonts would render their own name as pictographs
              <span style={{ fontFamily: isSymbolFontFamily(f) ? undefined : cssFontFamily(f) }}>
                {f}
              </span>
            ),
          })),
          ...(value && !fontFamilies.includes(value) && !systemFontFamilies.includes(value)
            ? [{ value, label: value } as DropdownOption]
            : []),
        ]}
        onPick={onPick}
      />
    </label>
  )

  const check = (label: string, checked: boolean, onChange: (v: boolean) => void) => (
    <label className="font-check">
      <input type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} />
      {label}
    </label>
  )

  const byInput = (value: number, onChange: (v: number) => void, disabled: boolean) => (
    <label className="font-by">
      {t('appCharBy')}
      <input
        type="number"
        step="0.5"
        min="0"
        value={value}
        disabled={disabled}
        onChange={(e) => onChange(Math.max(0, Number(e.target.value) || 0))}
      />
    </label>
  )

  const spacingPt =
    adv.spacing === 'normal' ? 0 : adv.spacingPt * (adv.spacing === 'condensed' ? -1 : 1)
  const scaleEm = adv.scalePct !== 100 ? charScaleEm(t('appFontPreviewSample'), adv.scalePct) : 0
  const positionPt =
    adv.position === 'normal' ? 0 : adv.positionPt * (adv.position === 'lowered' ? -1 : 1)
  const decorations = [
    underline ? 'underline' : '',
    strike && dstrike ? 'line-through double' : strike || dstrike ? 'line-through' : '',
    hidden ? 'underline dotted' : '',
  ]
    .filter(Boolean)
    .join(' ')

  const fontTab = (
    <>
      <div className="font-dialog-row">
        {fontPicker(fontLatin, field('fontLatin', setFontLatin), t('ribbonFontLatin'))}
        {fontPicker(fontEastAsia, field('fontEastAsia', setFontEastAsia), t('ribbonFontEastAsia'))}
        <label>
          {t('appFontStyleLabel')}
          <Dropdown
            value={style}
            ariaLabel={t('appFontStyleLabel')}
            options={FONT_STYLES.map((s) => ({ value: s.key, label: t(s.nameKey) }))}
            onPick={field('style', setStyle)}
          />
        </label>
        <label>
          {t('appFontSizeLabel')}
          <Dropdown
            value={fontSizeLabel(size, lang)}
            ariaLabel={t('appFontSizeLabel')}
            options={[
              ...(sizeOptions.some((o) => o.name === fontSizeLabel(size, lang))
                ? []
                : [fontSizeLabel(size, lang)]),
              ...sizeOptions.map((o) => o.name),
            ].map((s) => ({ value: s, label: s }))}
            onPick={(v) => {
              const pt = parseFontSize(v, lang)
              if (pt !== null) field('size', setSize)(pt)
            }}
          />
        </label>
      </div>
      <div className="font-dialog-row">
        <label>
          {t('appFontColor')}
          <input
            type="color"
            className="font-color-input"
            value={color}
            onChange={(e) => field('color', setColor)(e.target.value)}
          />
        </label>
        {check(t('appUnderline'), underline, field('underline', setUnderline))}
        {check(t('appStrikethrough'), strike, field('strike', setStrike))}
        {check(t('appSuperscript'), vertAlign === 'superscript', (v) =>
          field('vertAlign', setVertAlign)(v ? 'superscript' : ''),
        )}
        {check(t('appSubscript'), vertAlign === 'subscript', (v) =>
          field('vertAlign', setVertAlign)(v ? 'subscript' : ''),
        )}
      </div>
      <div className="font-dialog-row font-effects" aria-label={t('appFontEffects')}>
        {check(t('appDoubleStrikethrough'), dstrike, (v) => {
          setDstrike(v)
          touch('dstrike')
        })}
        {check(t('appSmallCaps'), caps === 'small', (v) => {
          setCaps(v ? 'small' : '')
          touch('caps')
        })}
        {check(t('appAllCaps'), caps === 'all', (v) => {
          setCaps(v ? 'all' : '')
          touch('caps')
        })}
        {check(t('appHiddenText'), hidden, (v) => {
          setHidden(v)
          touch('hidden')
        })}
      </div>
    </>
  )

  const advancedTab = (
    <>
      <div className="font-dialog-row">
        <label>
          {t('appCharScale')}
          <input
            type="number"
            list="font-scale-presets"
            min="1"
            max="600"
            value={adv.scalePct}
            onChange={(e) =>
              setAdvField(
                'scale',
                'scalePct',
                Math.min(600, Math.max(1, Number(e.target.value) || 100)),
              )
            }
          />
          <datalist id="font-scale-presets">
            {SCALE_PRESETS.map((p) => (
              <option key={p} value={p} />
            ))}
          </datalist>
        </label>
      </div>
      <div className="font-dialog-row">
        <label>
          {t('appCharSpacing')}
          <Dropdown
            value={adv.spacing}
            ariaLabel={t('appCharSpacing')}
            options={[
              { value: 'normal', label: t('appCharSpacingNormal') },
              { value: 'expanded', label: t('appCharSpacingExpanded') },
              { value: 'condensed', label: t('appCharSpacingCondensed') },
            ]}
            onPick={(v) => setAdvField('spacing', 'spacing', v)}
          />
        </label>
        {byInput(
          adv.spacingPt,
          (v) => setAdvField('spacing', 'spacingPt', v),
          adv.spacing === 'normal',
        )}
      </div>
      <div className="font-dialog-row">
        <label>
          {t('appCharPosition')}
          <Dropdown
            value={adv.position}
            ariaLabel={t('appCharPosition')}
            options={[
              { value: 'normal', label: t('appCharSpacingNormal') },
              { value: 'raised', label: t('appCharPositionRaised') },
              { value: 'lowered', label: t('appCharPositionLowered') },
            ]}
            onPick={(v) => setAdvField('position', 'position', v)}
          />
        </label>
        {byInput(
          adv.positionPt,
          (v) => setAdvField('position', 'positionPt', v),
          adv.position === 'normal',
        )}
      </div>
      <div className="font-dialog-row font-kern">
        {check(t('appKerningFor'), adv.kern, (v) => setAdvField('kern', 'kern', v))}
        <input
          type="number"
          step="0.5"
          min="0.5"
          aria-label={t('appKerningFor')}
          value={adv.kernPt}
          disabled={!adv.kern}
          onChange={(e) =>
            setAdvField('kern', 'kernPt', Math.max(0.5, Number(e.target.value) || 0.5))
          }
        />
        <span>{t('appKerningPointsAbove')}</span>
      </div>
    </>
  )

  return (
    <div
      className="modal-backdrop"
      ref={modalKeys.ref}
      onKeyDown={modalKeys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div className="modal font-dialog">
        <h2>{t('appFontDialogTitle')}</h2>
        <div className="dlg-tabs" role="tablist">
          {(['font', 'advanced'] as Tab[]).map((k) => (
            <button
              key={k}
              type="button"
              role="tab"
              aria-selected={tab === k}
              className={tab === k ? 'active' : undefined}
              onClick={() => setTab(k)}
            >
              {t(k === 'font' ? 'appFontTabFont' : 'appFontTabAdvanced')}
            </button>
          ))}
        </div>
        <div role="tabpanel">{tab === 'font' ? fontTab : advancedTab}</div>
        <div
          className="font-preview"
          style={{
            fontFamily:
              [fontLatin, fontEastAsia]
                .filter(Boolean)
                .map((f) => cssFontFamily(f))
                .join(', ') || undefined,
            fontSize: `${Math.min(size, 28)}pt`,
            fontWeight: style === 'bold' || style === 'boldItalic' ? 600 : 400,
            fontStyle: style === 'italic' || style === 'boldItalic' ? 'italic' : 'normal',
            color,
            textDecoration: decorations || undefined,
            textTransform: caps === 'all' ? 'uppercase' : undefined,
            fontVariantCaps: caps === 'small' ? 'small-caps' : undefined,
            opacity: hidden ? 0.5 : undefined,
            letterSpacing: spacingPt || scaleEm ? `calc(${spacingPt}pt + ${scaleEm}em)` : undefined,
            fontKerning: touched.has('kern') ? (adv.kern ? 'normal' : 'none') : undefined,
          }}
        >
          <span style={{ verticalAlign: positionPt ? `${positionPt}pt` : undefined }}>
            {t('appFontPreviewSample')}
          </span>
        </div>
        <div className="modal-actions">
          <button className="btn-ghost" onClick={onClose}>
            {t('appCancel')}
          </button>
          <button className="btn-primary" onClick={apply}>
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}

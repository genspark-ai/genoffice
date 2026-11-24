import { useState, type CSSProperties } from 'react'
import type { TextFill } from '@genoffice/docx-engine'
import {
  FX_COLORS,
  GLOW_SIZES_PT,
  GRADIENT_SPREADS,
  OUTLINE_WIDTHS_PT,
  REFLECTION_PRESETS,
  SHADOW_PRESETS,
  TEXT_FX_PRESETS,
  glowOf,
  gradientFillOf,
  gradientSpreadOf,
  sameTextFx,
  textFxDecls,
  type GradientSpread,
  type TextFx,
} from '../editor/text-fx'
import { declStyle } from '../editor/text-effects'
import { useI18n, type StringKey } from '../i18n/locale'
import { useModalKeys } from './modal-keys'

const SHADOW_LABELS: Record<(typeof SHADOW_PRESETS)[number]['id'], StringKey> = {
  right: 'ribbonTfxShadowRight',
  below: 'ribbonTfxShadowBelow',
  left: 'ribbonTfxShadowLeft',
}

const REFLECTION_LABELS: Record<(typeof REFLECTION_PRESETS)[number]['id'], StringKey> = {
  tight: 'ribbonTfxReflectTight',
  half: 'ribbonTfxReflectHalf',
  full: 'ribbonTfxReflectFull',
}

const GRADIENT_LABELS: Record<GradientSpread, StringKey> = {
  down: 'ribbonTfxGradientDown',
  right: 'ribbonTfxGradientRight',
  diagonal: 'ribbonTfxGradientDiagonal',
  center: 'ribbonTfxGradientCenter',
}

const AUTO_BASE = '000000'

/** CSS declarations ("name:value") as a React style object */
export const styleOf = (decls: string[]): CSSProperties => declStyle(decls)

/** Word's Format > Text Effects: presets, a fill, an outline, a shadow, a reflection and a glow on the selected text */
export function TextEffectsDialog({
  fx,
  color,
  shown,
  onApply,
  onClose,
}: {
  /** the effects the selected text has now */
  fx: TextFx
  /** the colour it is filled with now, hex without '#'; null for the automatic colour */
  color: string | null
  /** the fill it shows where that is more than a colour: a gradient, a transparent colour */
  shown: TextFill | null
  /** `color` is the fill that was picked, alone or with a preset; `fill` a gradient put on (null: taken off) */
  onApply: (fx: TextFx, color?: string, fill?: TextFill | null) => void
  onClose: () => void
}) {
  const { t } = useI18n()
  const keys = useModalKeys(onClose)
  const [next, setNextFx] = useState<TextFx>(fx)
  const [fill, setFill] = useState(color)
  const opened = shown?.kind === 'gradient' ? shown : null
  const [gradient, setGradient] = useState<TextFill | null>(opened)
  const spread = gradientSpreadOf(gradient)
  const gradientChanged = JSON.stringify(gradient) !== JSON.stringify(opened)
  // Word's Text Fill > Transparency: how much of a solid fill lets the page through
  const openedTransparency =
    shown?.kind === 'solid' ? Math.round((1 - (shown.alpha ?? 1)) * 100) : 0
  const [transparency, setTransparency] = useState(openedTransparency)
  const seeThrough: TextFill | null =
    !gradient && transparency > 0
      ? { kind: 'solid', color: fill ?? AUTO_BASE, alpha: (100 - transparency) / 100 }
      : null
  const seeThroughChanged =
    !gradient &&
    (transparency !== openedTransparency || (transparency > 0 && (fill !== color || !!opened)))
  const pickFill = (hex: string) => {
    setFill(hex)
    // one of our gradients follows the colour; any other gives way to it
    setGradient(spread ? gradientFillOf(hex, spread) : null)
  }
  const [preset, setPreset] = useState<(typeof TEXT_FX_PRESETS)[number] | null>(null)
  const setNext = (value: TextFx) => {
    setPreset(null)
    setNextFx(value)
  }
  const picked = preset || fill !== color ? (fill ?? undefined) : undefined
  const outlineColor = next.outline?.color ?? '000000'
  const glowColor = next.glow?.color ?? '4472C4'
  const choice = (active: boolean, label: string, onClick: () => void, key?: string) => (
    <button
      key={key ?? label}
      type="button"
      role="radio"
      aria-checked={active}
      className={`text-fx-choice${active ? ' active' : ''}`}
      onClick={onClick}
    >
      {label}
    </button>
  )
  const swatches = (
    label: string,
    current: string | null,
    pick: (color: string) => void,
    colors: readonly string[] = FX_COLORS,
  ) => (
    <div className="text-fx-colors" role="radiogroup" aria-label={label}>
      {colors.map((color) => (
        <button
          key={color}
          type="button"
          role="radio"
          aria-checked={current === color}
          aria-label={`#${color}`}
          className={`text-fx-swatch${current === color ? ' active' : ''}`}
          // a document colour, not chrome
          style={{ background: `#${color}` }}
          onClick={() => pick(color)}
        />
      ))}
    </div>
  )
  return (
    <div
      className="modal-backdrop"
      ref={keys.ref}
      onKeyDown={keys.onKeyDown}
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <form
        className="modal gs-form tco-dialog hlc-dialog text-fx-dialog"
        noValidate
        onSubmit={(e) => {
          e.preventDefault()
          onClose()
          if (picked || gradientChanged || seeThroughChanged || !sameTextFx(next, fx))
            onApply(
              next,
              // a see-through fill lies over a colour: the automatic colour becomes the one it was drawn in
              picked ?? (seeThrough && !fill ? AUTO_BASE : undefined),
              seeThroughChanged ? seeThrough : gradientChanged ? gradient : undefined,
            )
        }}
      >
        <h2>{t('ribbonTextEffects')}</h2>
        <fieldset className="text-fx-group" data-fx="presets">
          <legend>{t('ribbonColPresets')}</legend>
          <div className="text-fx-choices" role="radiogroup" aria-label={t('ribbonColPresets')}>
            {TEXT_FX_PRESETS.map((item, index) => (
              <button
                key={item.id}
                type="button"
                role="radio"
                data-preset={item.id}
                aria-checked={preset === item}
                aria-label={`${t('ribbonColPresets')} ${index + 1}`}
                className={`text-fx-preset${preset === item ? ' active' : ''}`}
                onClick={() => {
                  setNextFx(item.fx)
                  setFill(item.color)
                  setGradient(null)
                  setTransparency(0)
                  setPreset(item)
                }}
              >
                {/* a document colour, not chrome */}
                <span style={{ ...styleOf(textFxDecls(item.fx)), color: `#${item.color}` }}>A</span>
              </button>
            ))}
          </div>
        </fieldset>
        <fieldset className="text-fx-group" data-fx="fill">
          <legend>{t('ribbonTfxFill')}</legend>
          {swatches(
            t('ribbonTfxFill'),
            fill,
            (hex) => {
              setPreset(null)
              pickFill(hex)
            },
            // a colour from the file that the palette does not hold is offered beside it
            color && !FX_COLORS.some((hex) => hex === color) ? [color, ...FX_COLORS] : FX_COLORS,
          )}
          <div className="text-fx-choices" role="radiogroup" aria-label={t('ribbonTfxGradient')}>
            <span className="text-fx-label">{t('ribbonTfxGradient')}</span>
            {choice(!gradient, t('ribbonNone'), () => setGradient(null), 'none')}
            {GRADIENT_SPREADS.map((item) =>
              choice(
                spread === item,
                t(GRADIENT_LABELS[item]),
                () => {
                  setPreset(null)
                  setFill(fill ?? AUTO_BASE)
                  setGradient(gradientFillOf(fill ?? AUTO_BASE, item))
                },
                item,
              ),
            )}
          </div>
          {!gradient && (
            <div className="text-fx-choices text-fx-transparency">
              <span className="text-fx-label">{t('ribbonWatermarkTransparency')}</span>
              <input
                type="range"
                min={0}
                max={100}
                value={transparency}
                onChange={(e) => setTransparency(Number(e.target.value))}
                aria-label={t('ribbonWatermarkTransparency')}
              />
              <input
                type="number"
                className="wm-pct"
                name="fillTransparency"
                min={0}
                max={100}
                value={transparency}
                onChange={(e) =>
                  setTransparency(
                    Math.max(0, Math.min(100, Math.round(Number(e.target.value)) || 0)),
                  )
                }
                aria-label={t('ribbonWatermarkTransparency')}
              />
              <span className="text-fx-label">%</span>
            </div>
          )}
        </fieldset>
        <fieldset className="text-fx-group" data-fx="outline">
          <legend>{t('ribbonTfxOutline')}</legend>
          <div className="text-fx-choices" role="radiogroup" aria-label={t('ribbonPbWidth')}>
            {choice(!next.outline, t('ribbonNone'), () => setNext({ ...next, outline: null }))}
            {OUTLINE_WIDTHS_PT.map((widthPt) =>
              choice(next.outline?.widthPt === widthPt, `${widthPt} pt`, () =>
                setNext({ ...next, outline: { color: outlineColor, widthPt } }),
              ),
            )}
          </div>
          {next.outline &&
            swatches(t('ribbonPbColor'), outlineColor, (color) =>
              setNext({ ...next, outline: { ...next.outline!, color } }),
            )}
        </fieldset>
        <fieldset className="text-fx-group" data-fx="shadow">
          <legend>{t('ribbonPbShadow')}</legend>
          <div className="text-fx-choices" role="radiogroup" aria-label={t('ribbonPbShadow')}>
            {choice(!next.shadow, t('ribbonNone'), () => setNext({ ...next, shadow: null }))}
            {SHADOW_PRESETS.map((preset) =>
              choice(
                next.shadow?.dirDeg === preset.shadow.dirDeg,
                t(SHADOW_LABELS[preset.id]),
                () => setNext({ ...next, shadow: preset.shadow }),
                preset.id,
              ),
            )}
          </div>
        </fieldset>
        <fieldset className="text-fx-group" data-fx="reflection">
          <legend>{t('ribbonTfxReflection')}</legend>
          <div className="text-fx-choices" role="radiogroup" aria-label={t('ribbonTfxReflection')}>
            {choice(!next.reflection, t('ribbonNone'), () =>
              setNext({ ...next, reflection: null }),
            )}
            {REFLECTION_PRESETS.map((item) =>
              choice(
                next.reflection?.endPos === item.reflection.endPos,
                t(REFLECTION_LABELS[item.id]),
                () => setNext({ ...next, reflection: item.reflection }),
                item.id,
              ),
            )}
          </div>
        </fieldset>
        <fieldset className="text-fx-group" data-fx="glow">
          <legend>{t('ribbonTfxGlow')}</legend>
          <div className="text-fx-choices" role="radiogroup" aria-label={t('appFontSizeLabel')}>
            {choice(!next.glow, t('ribbonNone'), () => setNext({ ...next, glow: null }))}
            {GLOW_SIZES_PT.map((radiusPt) =>
              choice(next.glow?.radiusPt === radiusPt, `${radiusPt} pt`, () =>
                setNext({ ...next, glow: glowOf(glowColor, radiusPt) }),
              ),
            )}
          </div>
          {next.glow &&
            swatches(t('ribbonPbColor'), glowColor, (color) =>
              setNext({ ...next, glow: glowOf(color, next.glow!.radiusPt) }),
            )}
        </fieldset>
        <div className="text-fx-preview" aria-label={t('ribbonPbPreview')}>
          <span
            style={{
              ...(fill ? { color: `#${fill}` } : {}),
              ...styleOf(
                textFxDecls(
                  next,
                  gradient ??
                    seeThrough ??
                    (fill === color && shown?.kind === 'none' ? shown : null),
                ),
              ),
            }}
          >
            AaBbCc
          </span>
        </div>
        <div className="modal-actions">
          <div className="tco-actions-right">
            <button type="button" onClick={onClose}>
              {t('ribbonCancel')}
            </button>
            <button type="submit" className="primary">
              {t('ribbonOk')}
            </button>
          </div>
        </div>
      </form>
    </div>
  )
}

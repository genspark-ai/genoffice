import { useEffect, useRef, useState } from 'react'
import type { MutableRefObject } from 'react'
import { useI18n } from '../i18n/locale'
import type { StringKey } from '../i18n/locale'
import type { ComputedSnapshot } from '../preview/inspector-protocol'
import type { StyleEdit } from './FloatToolbar'
import { ColorField } from './ColorField'
import { backgroundPickStyles, backgroundSwatch } from '../document/background-style'

interface Props {
  tag: string
  computed: ComputedSnapshot
  /** current single text run, null when the element has none (or several) */
  textRun: string | null
  onStyle: StyleEdit
  onText: (text: string) => void
  /** attribute write (alt text on images) */
  onAttr: (name: string, value: string | null) => void
  onCustomCss: (css: string) => void
  /** the host calls this before structural ops / selection changes so open drafts land on the right element */
  draftRef: MutableRefObject<(() => void) | null>
  /** live previews not yet written to the source */
  pending: boolean
  onRevert: () => void
  onClose: () => void
}

type NumberField = 'fontSize' | 'width' | 'height' | 'borderRadius' | 'padding'

const PX_FIELDS: Record<NumberField, { prop: string; label: StringKey; min: number }> = {
  fontSize: { prop: 'font-size', label: 'fontSize', min: 1 },
  width: { prop: 'width', label: 'width', min: 0 },
  height: { prop: 'height', label: 'height', min: 0 },
  borderRadius: { prop: 'border-radius', label: 'radius', min: 0 },
  padding: { prop: 'padding', label: 'padding', min: 0 },
}

function PxInput({
  field,
  value,
  onStyle,
}: {
  field: NumberField
  value: string
  onStyle: StyleEdit
}) {
  const { t } = useI18n()
  const def = PX_FIELDS[field]
  return (
    <label>
      <span>{t(def.label)}</span>
      <input
        type="number"
        min={def.min}
        value={value}
        onChange={(e) => {
          const v = e.target.value
          if (v !== '' && Number(v) >= def.min) onStyle({ [def.prop]: `${v}px` })
        }}
      />
    </label>
  )
}

/**
 * A text field draft committed on blur / Enter, and otherwise by the host (before ops, on
 * selection change) or when the field unmounts: toolbar buttons swallow mousedown, so blur
 * alone is not enough. Idempotent so the several flush points write at most once.
 */
function useDraft(base: string, write: (value: string) => void) {
  const [value, setValue] = useState(base)
  const committedRef = useRef(base)
  const valueRef = useRef(value)
  valueRef.current = value
  useEffect(() => {
    // the frame re-reports what we just wrote; keep keystrokes typed since that commit
    const untouched = valueRef.current === committedRef.current
    committedRef.current = base
    if (untouched) setValue(base)
  }, [base])
  const commitRef = useRef<() => void>(() => {})
  commitRef.current = () => {
    if (value !== committedRef.current) {
      committedRef.current = value
      write(value)
    }
  }
  const commit = () => commitRef.current()
  return { value, setValue, commit, commitRef }
}

const FITS = ['fill', 'contain', 'cover', 'none'] as const
const FIT_LABEL: Record<(typeof FITS)[number], StringKey> = {
  fill: 'imageFitFill',
  contain: 'imageFitContain',
  cover: 'imageFitCover',
  none: 'imageFitNone',
}

/** <img>: alt text (attribute) and object-fit; typography has nothing to say about a replaced element */
function ImageSection(p: Props & { altCommitRef: MutableRefObject<() => void> }) {
  const { t } = useI18n()
  const image = p.computed.image
  const {
    value: alt,
    setValue: setAlt,
    commit: commitAlt,
    commitRef,
  } = useDraft(image?.alt ?? '', (v) => p.onAttr('alt', v))
  p.altCommitRef.current = () => commitRef.current()
  const fit = (FITS as readonly string[]).includes(p.computed.objectFit)
    ? p.computed.objectFit
    : 'fill'
  return (
    <>
      <div className="hx-panel-section">{t('imageSection')}</div>
      <label>
        <span>{t('imageAlt')}</span>
        <input
          type="text"
          value={alt}
          onChange={(e) => setAlt(e.target.value)}
          onBlur={commitAlt}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault()
              commitAlt()
            }
          }}
        />
      </label>
      <div className="hx-panel-row">
        <label>
          <span>{t('imageFit')}</span>
          <select value={fit} onChange={(e) => p.onStyle({ 'object-fit': e.target.value })}>
            {FITS.map((f) => (
              <option key={f} value={f}>
                {t(FIT_LABEL[f])}
              </option>
            ))}
          </select>
        </label>
        {image && image.naturalWidth > 0 && (
          <label>
            <span>{t('imageNatural')}</span>
            <input type="text" readOnly value={`${image.naturalWidth} × ${image.naturalHeight}`} />
          </label>
        )}
      </div>
    </>
  )
}

/** right-hand style inspector for the selected element: every change previews live and lands in the source as one set_style op */
export function StylePanel(p: Props) {
  const { t } = useI18n()
  const isImage = p.tag === 'img'
  const {
    value: text,
    setValue: setText,
    commit: commitText,
    commitRef: textCommitRef,
  } = useDraft(p.textRun ?? '', (v) => {
    if (p.textRun !== null) p.onText(v)
  })
  const altCommitRef = useRef<() => void>(() => {})
  const [css, setCss] = useState('')
  const flushDrafts = () => {
    textCommitRef.current()
    altCommitRef.current()
  }
  p.draftRef.current = flushDrafts
  const draftRef = p.draftRef
  useEffect(
    () => () => {
      draftRef.current?.()
      draftRef.current = null
    },
    [draftRef],
  )
  const commitCss = () => {
    if (css.trim()) {
      p.onCustomCss(css.trim())
      setCss('')
    }
  }

  return (
    <aside className="hx-panel" aria-label={t('stylePanel')}>
      <div className="hx-panel-head">
        <span className="hx-panel-tag">&lt;{p.tag}&gt;</span>
        {p.pending && (
          <button type="button" className="hx-panel-revert" onClick={p.onRevert}>
            {t('revert')}
          </button>
        )}
        <button
          type="button"
          className="hx-panel-close"
          aria-label={t('stylePanelClose')}
          onClick={p.onClose}
        >
          ×
        </button>
      </div>

      {isImage ? (
        <ImageSection {...p} altCommitRef={altCommitRef} />
      ) : (
        <>
          <div className="hx-panel-section">{t('typography')}</div>
          <div className="hx-panel-row">
            <PxInput field="fontSize" value={p.computed.fontSize} onStyle={p.onStyle} />
            <label>
              <span>{t('fontWeight')}</span>
              <select
                value={p.computed.fontWeight}
                onChange={(e) => p.onStyle({ 'font-weight': e.target.value })}
              >
                {['300', '400', '500', '600', '700', '800'].map((w) => (
                  <option key={w} value={w}>
                    {w}
                  </option>
                ))}
                {!['300', '400', '500', '600', '700', '800'].includes(p.computed.fontWeight) && (
                  <option value={p.computed.fontWeight}>{p.computed.fontWeight}</option>
                )}
              </select>
            </label>
          </div>
          <div className="hx-panel-row">
            <ColorField
              label={t('color')}
              value={p.computed.color}
              onPick={(hex) => hex && p.onStyle({ color: hex })}
            />
            <label>
              <span>{t('align')}</span>
              <select
                value={
                  ['left', 'center', 'right', 'justify'].includes(p.computed.textAlign)
                    ? p.computed.textAlign
                    : 'left'
                }
                onChange={(e) => p.onStyle({ 'text-align': e.target.value })}
              >
                <option value="left">{t('alignLeft')}</option>
                <option value="center">{t('alignCenter')}</option>
                <option value="right">{t('alignRight')}</option>
              </select>
            </label>
          </div>
        </>
      )}

      <div className="hx-panel-section">{t('sizeSection')}</div>
      <div className="hx-panel-row">
        <PxInput field="width" value={p.computed.width} onStyle={p.onStyle} />
        <PxInput field="height" value={p.computed.height} onStyle={p.onStyle} />
      </div>
      {isImage && (
        <button
          type="button"
          className="hx-panel-action"
          onClick={() => p.onStyle({ width: null, height: null })}
        >
          {t('imageResetSize')}
        </button>
      )}

      <div className="hx-panel-section">{t('boxSection')}</div>
      <div className="hx-panel-row">
        {!isImage && (
          <ColorField
            label={t('background')}
            value={p.computed.background}
            swatch={backgroundSwatch(p.computed.background, p.computed.backgroundImage)}
            none
            onPick={(hex) => p.onStyle(backgroundPickStyles(hex, p.computed.backgroundImage))}
          />
        )}
        <label>
          <span>{t('opacity')}</span>
          <input
            type="number"
            min={0}
            max={1}
            step={0.05}
            value={p.computed.opacity}
            onChange={(e) => {
              const v = Number(e.target.value)
              if (e.target.value !== '' && v >= 0 && v <= 1) p.onStyle({ opacity: String(v) })
            }}
          />
        </label>
      </div>
      <div className="hx-panel-row">
        <PxInput field="borderRadius" value={p.computed.borderRadius} onStyle={p.onStyle} />
        <PxInput field="padding" value={p.computed.padding} onStyle={p.onStyle} />
      </div>

      {p.textRun !== null && (
        <>
          <div className="hx-panel-section">{t('textLabel')}</div>
          <label>
            <textarea
              rows={3}
              value={text}
              onChange={(e) => setText(e.target.value)}
              onBlur={commitText}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
                  e.preventDefault()
                  commitText()
                }
              }}
            />
          </label>
        </>
      )}

      <label>
        <span>{t('customCss')}</span>
        <input
          type="text"
          placeholder="letter-spacing: 2px; box-shadow: …"
          value={css}
          onChange={(e) => setCss(e.target.value)}
          onBlur={commitCss}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault()
              commitCss()
            }
          }}
        />
      </label>
    </aside>
  )
}

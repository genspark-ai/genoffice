import { useState } from 'react'
import type { CSSProperties, ReactElement } from 'react'
import { useI18n } from '../i18n/locale'
import type { StringKey } from '../i18n/locale'
import type { Brief, BriefSection, BriefStyle } from '../document/brief'
import type { BriefDecision } from '../ai/tools'

const DENSITIES = ['compact', 'regular', 'airy'] as const
const PALETTE_KEYS = ['primary', 'accent', 'bg', 'surface', 'text'] as const
const TOKEN_KEYS = ['radius', 'spacing', 'shadow'] as const
type PaletteKey = (typeof PALETTE_KEYS)[number]
type TokenKey = (typeof TOKEN_KEYS)[number]
const PALETTE_LABEL: Record<PaletteKey, StringKey> = {
  primary: 'briefPrimary',
  accent: 'briefAccent',
  bg: 'briefBg',
  surface: 'briefSurface',
  text: 'briefText',
}
const TOKEN_LABEL: Record<TokenKey, StringKey> = {
  radius: 'briefRadius',
  spacing: 'briefSpacing',
  shadow: 'briefShadow',
}

function isHex(v: string | undefined): v is string {
  return !!v && /^#[0-9a-fA-F]{6}$/.test(v)
}

function cloneStyle(s: BriefStyle): BriefStyle {
  return {
    ...s,
    palette: { ...s.palette },
    typography: { ...s.typography },
    tokens: { ...s.tokens },
  }
}

/** Miniature of one direction rendered from its own palette / fonts / tokens: document colors, not chrome. */
function StyleTile({
  style,
  heading,
  cta,
  selected,
  onSelect,
}: {
  style: BriefStyle
  heading: string
  cta: string
  selected: boolean
  onSelect: () => void
}): ReactElement {
  const p = style.palette
  const bg = isHex(p.bg) ? p.bg : '#ffffff'
  const text = isHex(p.text) ? p.text : '#1f2328'
  const primary = isHex(p.primary) ? p.primary : text
  const accent = isHex(p.accent) ? p.accent : primary
  const surface = isHex(p.surface) ? p.surface : bg
  const radius = style.tokens?.radius || '8px'
  const shadow = style.tokens?.shadow || 'none'
  const page: CSSProperties = { background: bg, color: text, fontFamily: style.typography.body }
  const card: CSSProperties = {
    background: surface,
    borderRadius: radius,
    boxShadow: shadow,
    border: surface === bg ? `1px solid ${text}22` : 'none',
  }
  return (
    <button
      type="button"
      className="brief-tile"
      aria-pressed={selected}
      onClick={onSelect}
      title={style.tone}
    >
      <div className="brief-tile-page" style={page}>
        <div
          className="brief-tile-heading"
          style={{ fontFamily: style.typography.heading, color: primary }}
        >
          {heading}
        </div>
        <div className="brief-tile-body">{style.tone}</div>
        <div className="brief-tile-row">
          <span
            className="brief-tile-cta"
            style={{ background: accent, color: bg, borderRadius: radius }}
          >
            {cta}
          </span>
          <span className="brief-tile-card" style={card}>
            <i style={{ background: primary }} />
            <i style={{ background: `${text}55` }} />
          </span>
        </div>
      </div>
      <div className="brief-tile-name">{style.name || style.tone}</div>
    </button>
  )
}

/**
 * Editable proposal shown while plan_page waits. The style directions are
 * tiles; the fields below edit the selected one. Every field the user touches
 * is recorded so the model keeps it verbatim; "another brief" hands a note back.
 */
export function BriefCard({
  brief,
  onDecide,
}: {
  brief: Brief
  onDecide: (decision: BriefDecision) => void
}): ReactElement {
  const { t } = useI18n()
  const [styles, setStyles] = useState<BriefStyle[]>(() =>
    [brief.style, ...(brief.alternatives ?? [])].map(cloneStyle),
  )
  const [sel, setSel] = useState(0)
  const [draft, setDraft] = useState<Omit<Brief, 'style' | 'alternatives'>>(() => ({
    core_hook: brief.core_hook,
    sections: brief.sections.map((s) => ({ ...s })),
    meta: { ...brief.meta },
    version: 1,
  }))
  const [edited, setEdited] = useState<Set<string>>(new Set())
  const [redoNote, setRedoNote] = useState<string | null>(null)
  const style = styles[sel] ?? styles[0]!

  const touch = (field: string) => setEdited((prev) => new Set(prev).add(field))
  const patchStyle = (fn: (s: BriefStyle) => BriefStyle, field: string) => {
    setStyles((all) => all.map((s, i) => (i === sel ? fn(s) : s)))
    touch(field)
  }
  const setStyle = (patch: Partial<BriefStyle>, field: string) =>
    patchStyle((s) => ({ ...s, ...patch }), field)
  const setPalette = (key: PaletteKey, value: string) =>
    patchStyle((s) => ({ ...s, palette: { ...s.palette, [key]: value } }), `style.palette.${key}`)
  const setTypography = (key: 'heading' | 'body', value: string) =>
    patchStyle(
      (s) => ({ ...s, typography: { ...s.typography, [key]: value } }),
      `style.typography.${key}`,
    )
  const setToken = (key: TokenKey, value: string) =>
    patchStyle((s) => ({ ...s, tokens: { ...s.tokens, [key]: value } }), `style.tokens.${key}`)
  const pickDirection = (i: number) => {
    setSel(i)
    if (i !== 0) touch('style')
  }
  const setSection = (i: number, patch: Partial<BriefSection>) => {
    setDraft((d) => ({
      ...d,
      sections: d.sections.map((s, k) => (k === i ? { ...s, ...patch } : s)),
    }))
    touch(`sections[${i}]`)
  }
  const removeSection = (i: number) => {
    setDraft((d) => ({ ...d, sections: d.sections.filter((_, k) => k !== i) }))
    touch('sections')
  }
  const addSection = () => {
    setDraft((d) => ({ ...d, sections: [...d.sections, { title: '', brief: '' }] }))
    touch('sections')
  }
  const setMeta = (key: 'title' | 'audience' | 'length' | 'language', value: string) => {
    setDraft((d) => ({ ...d, meta: { ...d.meta, [key]: value } }))
    touch(`meta.${key}`)
  }

  const canGenerate =
    draft.core_hook.trim().length > 0 && draft.sections.some((s) => s.title.trim())

  const confirm = () =>
    onDecide({
      kind: 'confirmed',
      brief: {
        ...draft,
        style,
        sections: draft.sections.filter((s) => s.title.trim() || s.brief.trim()),
        user_edited: [...edited],
      },
    })

  const tileHeading = draft.meta?.title?.trim() || draft.core_hook
  const tileCta = draft.sections.find((s) => s.title.trim())?.title ?? '—'

  return (
    <div className="brief-card" role="form" aria-label={t('briefTitle')}>
      <div className="brief-head">
        <span>{t('briefTitle')}</span>
        <span className="brief-hint">{t('briefHint')}</span>
      </div>

      <div className="brief-field">
        <label htmlFor="brief-hook">{t('briefCoreHook')}</label>
        <textarea
          id="brief-hook"
          rows={2}
          value={draft.core_hook}
          onChange={(e) => {
            setDraft((d) => ({ ...d, core_hook: e.target.value }))
            touch('core_hook')
          }}
        />
      </div>

      <div className="brief-field">
        <label>{t('briefDirections')}</label>
        <div className="brief-directions" role="radiogroup" aria-label={t('briefDirections')}>
          {styles.map((s, i) => (
            <StyleTile
              key={i}
              style={s}
              heading={tileHeading}
              cta={tileCta}
              selected={i === sel}
              onSelect={() => pickDirection(i)}
            />
          ))}
        </div>
      </div>

      <div className="brief-row">
        <div className="brief-field">
          <label htmlFor="brief-tone">{t('briefTone')}</label>
          <input
            id="brief-tone"
            type="text"
            value={style.tone}
            onChange={(e) => setStyle({ tone: e.target.value }, 'style.tone')}
          />
        </div>
        <div className="brief-field">
          <label htmlFor="brief-density">{t('briefDensity')}</label>
          <select
            id="brief-density"
            value={style.density ?? 'regular'}
            onChange={(e) => setStyle({ density: e.target.value }, 'style.density')}
          >
            {DENSITIES.map((d) => (
              <option key={d} value={d}>
                {d}
              </option>
            ))}
          </select>
        </div>
      </div>

      <div className="brief-field">
        <label>{t('briefPalette')}</label>
        <div className="brief-palette">
          {PALETTE_KEYS.map((key) => {
            const value = style.palette[key]
            return (
              <label key={key} className="brief-swatch">
                <input
                  type="color"
                  value={
                    isHex(value) ? value : key === 'bg' || key === 'surface' ? '#ffffff' : '#333333'
                  }
                  aria-label={t(PALETTE_LABEL[key])}
                  onChange={(e) => setPalette(key, e.target.value)}
                />
                {t(PALETTE_LABEL[key])}
              </label>
            )
          })}
        </div>
      </div>

      <div className="brief-row">
        <div className="brief-field">
          <label htmlFor="brief-heading-font">{t('briefHeadingFont')}</label>
          <input
            id="brief-heading-font"
            type="text"
            value={style.typography.heading ?? ''}
            onChange={(e) => setTypography('heading', e.target.value)}
          />
        </div>
        <div className="brief-field">
          <label htmlFor="brief-body-font">{t('briefBodyFont')}</label>
          <input
            id="brief-body-font"
            type="text"
            value={style.typography.body ?? ''}
            onChange={(e) => setTypography('body', e.target.value)}
          />
        </div>
      </div>

      <div className="brief-row brief-row-3">
        {TOKEN_KEYS.map((key) => (
          <div key={key} className="brief-field">
            <label htmlFor={`brief-token-${key}`}>{t(TOKEN_LABEL[key])}</label>
            <input
              id={`brief-token-${key}`}
              type="text"
              value={style.tokens?.[key] ?? ''}
              onChange={(e) => setToken(key, e.target.value)}
            />
          </div>
        ))}
      </div>

      <div className="brief-row">
        <div className="brief-field">
          <label htmlFor="brief-layout">{t('briefLayout')}</label>
          <input
            id="brief-layout"
            type="text"
            value={style.layout ?? ''}
            onChange={(e) => setStyle({ layout: e.target.value }, 'style.layout')}
          />
        </div>
        <div className="brief-field">
          <label htmlFor="brief-docx">{t('briefDocxFriendly')}</label>
          <input
            id="brief-docx"
            type="checkbox"
            checked={style.docx_friendly === true}
            onChange={(e) => setStyle({ docx_friendly: e.target.checked }, 'style.docx_friendly')}
          />
        </div>
      </div>

      <div className="brief-field">
        <label>{t('briefSections')}</label>
        <div className="brief-sections">
          {draft.sections.map((s, i) => (
            <div key={i} className="brief-section">
              <input
                type="text"
                value={s.title}
                onChange={(e) => setSection(i, { title: e.target.value })}
                aria-label={t('briefSections')}
              />
              <button
                type="button"
                className="brief-section-remove"
                aria-label={t('briefRemoveSection')}
                title={t('briefRemoveSection')}
                onClick={() => removeSection(i)}
              >
                ×
              </button>
              <textarea
                rows={2}
                value={s.brief}
                placeholder={t('briefSectionBrief')}
                onChange={(e) => setSection(i, { brief: e.target.value })}
              />
            </div>
          ))}
          <button type="button" className="brief-add" onClick={addSection}>
            {t('briefAddSection')}
          </button>
        </div>
      </div>

      <div className="brief-row">
        <div className="brief-field">
          <label htmlFor="brief-audience">{t('briefAudience')}</label>
          <input
            id="brief-audience"
            type="text"
            value={draft.meta?.audience ?? ''}
            onChange={(e) => setMeta('audience', e.target.value)}
          />
        </div>
        <div className="brief-field">
          <label htmlFor="brief-length">{t('briefLength')}</label>
          <input
            id="brief-length"
            type="text"
            value={draft.meta?.length ?? ''}
            onChange={(e) => setMeta('length', e.target.value)}
          />
        </div>
      </div>

      <div className="brief-actions">
        {redoNote !== null ? (
          <>
            <input
              className="brief-redo-input"
              autoFocus
              value={redoNote}
              placeholder={t('briefRedoPlaceholder')}
              onChange={(e) => setRedoNote(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') onDecide({ kind: 'redo', note: redoNote.trim() })
                if (e.key === 'Escape') setRedoNote(null)
              }}
            />
            <button
              type="button"
              className="brief-btn"
              onClick={() => onDecide({ kind: 'redo', note: redoNote.trim() })}
            >
              {t('briefRedo')}
            </button>
          </>
        ) : (
          <>
            <button
              type="button"
              className="brief-btn"
              onClick={() => onDecide({ kind: 'cancelled' })}
            >
              {t('briefDismiss')}
            </button>
            <button type="button" className="brief-btn" onClick={() => setRedoNote('')}>
              {t('briefRedo')}
            </button>
            <button
              type="button"
              className="brief-btn primary"
              disabled={!canGenerate}
              onClick={confirm}
            >
              {t('briefGenerate')}
            </button>
          </>
        )}
      </div>
    </div>
  )
}

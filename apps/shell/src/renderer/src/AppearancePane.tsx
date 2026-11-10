import { useEffect, useState, useSyncExternalStore } from 'react'
import type { CSSProperties, ReactNode } from 'react'
import {
  DARK_PRESETS,
  LIGHT_PRESETS,
  contrastRatio,
  getAppearance,
  matchPreset,
  parseHex,
  resetAppearance,
  subscribeAppearance,
  updateAppearance,
} from './appearance'
import type { Contrast, ConversationSteps, ThemeColors, ThemePreset, UiFont } from './appearance'
import { useI18n } from './locale'
import type { StringKey } from './locale'
import type { UiTheme } from '../../shared/home-api'

// ── Settings › Appearance ───────────────────────────────────
// Chat · Appearance · Dark Theme · Light Theme · Interface. Every control writes
// through the appearance store, which persists, repaints <html> and notifies.

const FONT_OPTIONS: readonly { value: UiFont; label: string | StringKey }[] = [
  { value: 'google-sans', label: 'Google Sans' },
  { value: 'system', label: 'fontSystem' },
  { value: 'sans', label: 'fontSans' },
  { value: 'serif', label: 'fontSerif' },
  { value: 'mono', label: 'fontMono' },
]

function useAppearance() {
  return useSyncExternalStore(subscribeAppearance, getAppearance, getAppearance)
}

function Group({ title, children }: { title: string; children: ReactNode }) {
  return (
    <section className="ap-group">
      <h4 className="ap-group-title">{title}</h4>
      <div className="ap-card">{children}</div>
    </section>
  )
}

function Row({
  label,
  hint,
  children,
}: {
  label: string
  hint?: string
  children: ReactNode
}) {
  return (
    <div className="ap-row">
      <div className="ap-row-text">
        <div className="ap-row-label">{label}</div>
        {hint && <div className="ap-row-hint">{hint}</div>}
      </div>
      <div className="ap-row-control">{children}</div>
    </div>
  )
}

function Segmented<T extends string>({
  value,
  options,
  label,
  onChange,
}: {
  value: T
  options: readonly { value: T; label: ReactNode; aria?: string }[]
  label: string
  onChange: (v: T) => void
}) {
  return (
    <div className="ap-seg" role="radiogroup" aria-label={label}>
      {options.map((o) => (
        <button
          key={o.value}
          type="button"
          role="radio"
          aria-checked={value === o.value}
          aria-label={o.aria}
          className={`ap-seg-btn${value === o.value ? ' on' : ''}`}
          onClick={() => onChange(o.value)}
        >
          {o.label}
        </button>
      ))}
    </div>
  )
}

function SelectPill({
  id,
  value,
  text,
  onChange,
  children,
}: {
  id: string
  value: string
  text: string
  onChange: (v: string) => void
  children: ReactNode
}) {
  return (
    <span className="set-select-wrap ap-pill">
      <span className="set-select-text" aria-hidden="true">
        {text}
      </span>
      <select id={id} className="set-select" value={value} onChange={(e) => onChange(e.target.value)}>
        {children}
      </select>
    </span>
  )
}

/** swatch + hex pill. The native picker opens from the swatch; typing commits on every valid hex. */
function ColorField({
  label,
  value,
  onChange,
  invalidText,
  pickLabel,
}: {
  label: string
  value: string
  onChange: (hex: string) => void
  invalidText: string
  pickLabel: string
}) {
  const [text, setText] = useState(value.slice(1).toUpperCase())
  const [invalid, setInvalid] = useState(false)

  // the color changed from outside (preset switch, reset, native picker): show it. Typing a valid
  // hex emits the same color back, so this never fights the user's keystrokes.
  useEffect(() => {
    if (parseHex(text) !== value) {
      setText(value.slice(1).toUpperCase())
      setInvalid(false)
    }
    // only react to the color itself
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value])

  const commit = (raw: string): void => {
    const parsed = parseHex(raw)
    if (parsed) {
      setInvalid(false)
      if (parsed !== value) onChange(parsed)
    } else {
      setInvalid(raw.trim() !== '')
    }
  }

  return (
    <span className="ap-color" data-invalid={invalid || undefined}>
      <label className="ap-swatch" style={{ '--swatch': value } as CSSProperties} title={pickLabel}>
        <input
          type="color"
          value={value}
          aria-label={`${label}: ${pickLabel}`}
          onChange={(e) => {
            const parsed = parseHex(e.target.value)
            if (parsed) onChange(parsed)
          }}
        />
      </label>
      <span className="ap-hash" aria-hidden="true">
        #
      </span>
      <input
        className="ap-hex"
        value={text}
        maxLength={7}
        spellCheck={false}
        aria-label={`${label} (hex)`}
        aria-invalid={invalid}
        title={invalid ? invalidText : undefined}
        onChange={(e) => {
          const next = e.target.value.replace(/^#/, '')
          setText(next.toUpperCase())
          commit(next)
        }}
        onBlur={() => {
          // snap back to the last good value if the text never became valid
          if (parseHex(text) === null) {
            setText(value.slice(1).toUpperCase())
            setInvalid(false)
          }
        }}
      />
    </span>
  )
}

function ThemeGroup({
  kind,
  title,
}: {
  kind: 'dark' | 'light'
  title: string
}) {
  const { t } = useI18n()
  const appearance = useAppearance()
  const colors = appearance[kind]
  const presets: readonly ThemePreset[] = kind === 'dark' ? DARK_PRESETS : LIGHT_PRESETS
  const presetId = matchPreset(kind, colors)
  const presetName = presets.find((p) => p.id === presetId)?.name ?? t('apPresetCustom')

  const set = (patch: Partial<ThemeColors>): void => {
    updateAppearance({ [kind]: { ...colors, ...patch } })
  }

  const textLow = contrastRatio(colors.background, colors.foreground) < 4.5
  const accentLow = contrastRatio(colors.background, colors.accent) < 3

  return (
    <Group title={title}>
      <Row label={t('apPreset')}>
        <SelectPill
          id={`ap-preset-${kind}`}
          value={presetId}
          text={presetName}
          onChange={(id) => {
            const hit = presets.find((p) => p.id === id)
            if (hit)
              updateAppearance({
                [kind]: { background: hit.background, foreground: hit.foreground, accent: hit.accent },
              })
          }}
        >
          {presets.map((p) => (
            <option key={p.id} value={p.id}>
              {p.name}
            </option>
          ))}
          {presetId === 'custom' && (
            <option value="custom" disabled>
              {t('apPresetCustom')}
            </option>
          )}
        </SelectPill>
      </Row>
      <Row label={t('apBackground')}>
        <ColorField
          label={t('apBackground')}
          value={colors.background}
          onChange={(background) => set({ background })}
          invalidText={t('appearanceAccentInvalid')}
          pickLabel={t('apPickColor')}
        />
      </Row>
      <Row label={t('apForeground')}>
        <ColorField
          label={t('apForeground')}
          value={colors.foreground}
          onChange={(foreground) => set({ foreground })}
          invalidText={t('appearanceAccentInvalid')}
          pickLabel={t('apPickColor')}
        />
      </Row>
      <Row label={t('apAccent')}>
        <ColorField
          label={t('apAccent')}
          value={colors.accent}
          onChange={(accent) => set({ accent })}
          invalidText={t('appearanceAccentInvalid')}
          pickLabel={t('apPickColor')}
        />
      </Row>
      {(textLow || accentLow) && (
        <div className="ap-warn" role="status">
          {textLow ? t('apLowContrastText') : t('appearanceAccentLow')}
        </div>
      )}
    </Group>
  )
}

function ThemeIcon({ id }: { id: UiTheme }) {
  if (id === 'system')
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <rect x="1.8" y="2.6" width="12.4" height="8.4" rx="1.6" stroke="currentColor" strokeWidth="1.3" />
        <path d="M5.5 13.6h5M8 11v2.6" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
      </svg>
    )
  if (id === 'light')
    return (
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <circle cx="8" cy="8" r="2.8" stroke="currentColor" strokeWidth="1.3" />
        <path
          d="M8 1.6v1.6M8 12.8v1.6M1.6 8h1.6M12.8 8h1.6M3.5 3.5l1.1 1.1M11.4 11.4l1.1 1.1M12.5 3.5l-1.1 1.1M4.6 11.4l-1.1 1.1"
          stroke="currentColor"
          strokeWidth="1.3"
          strokeLinecap="round"
        />
      </svg>
    )
  return (
    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
      <path
        d="M13.4 9.6A5.8 5.8 0 0 1 6.4 2.6a5.8 5.8 0 1 0 7 7z"
        stroke="currentColor"
        strokeWidth="1.3"
        strokeLinejoin="round"
      />
    </svg>
  )
}

export function AppearancePane({
  theme,
  onTheme,
}: {
  theme: UiTheme
  onTheme: (theme: UiTheme) => void
}) {
  const { t } = useI18n()
  const appearance = useAppearance()

  const themeOptions = [
    { value: 'system', label: <ThemeIcon id="system" />, aria: t('themeSystem') },
    { value: 'light', label: <ThemeIcon id="light" />, aria: t('themeLight') },
    { value: 'dark', label: <ThemeIcon id="dark" />, aria: t('themeDark') },
  ] as const satisfies readonly { value: UiTheme; label: ReactNode; aria: string }[]

  const fontLabel = (o: (typeof FONT_OPTIONS)[number]): string =>
    o.label === 'Google Sans' ? o.label : t(o.label as StringKey)

  return (
    <>
      <h3 className="set-pane-title ap-title">{t('setSecAppearance')}</h3>
      <p className="ap-sub">{t('setAppearanceSub')}</p>

      <Group title={t('apGroupChat')}>
        <Row label={t('apConvSteps')} hint={t('apConvStepsHint')}>
          <Segmented<ConversationSteps>
            value={appearance.conversationSteps}
            label={t('apConvSteps')}
            options={[
              { value: 'default', label: t('apStepsDefault') },
              { value: 'verbose', label: t('apStepsVerbose') },
            ]}
            onChange={(conversationSteps) => updateAppearance({ conversationSteps })}
          />
        </Row>
      </Group>

      <Group title={t('apGroupAppearance')}>
        <Row label={t('theme')}>
          <Segmented<UiTheme>
            value={theme}
            label={t('theme')}
            options={themeOptions}
            onChange={onTheme}
          />
        </Row>
        <Row label={t('apContrast')}>
          <Segmented<Contrast>
            value={appearance.contrast}
            label={t('apContrast')}
            options={[
              { value: 'default', label: t('apContrastDefault') },
              { value: 'strong', label: t('apContrastStrong') },
            ]}
            onChange={(contrast) => updateAppearance({ contrast })}
          />
        </Row>
        <Row label={t('appearanceFont')}>
          <SelectPill
            id="ap-font"
            value={appearance.font}
            text={fontLabel(FONT_OPTIONS.find((o) => o.value === appearance.font) ?? FONT_OPTIONS[1]!)}
            onChange={(font) => updateAppearance({ font: font as UiFont })}
          >
            {FONT_OPTIONS.map((o) => (
              <option key={o.value} value={o.value}>
                {fontLabel(o)}
              </option>
            ))}
          </SelectPill>
        </Row>
      </Group>

      <ThemeGroup kind="dark" title={t('apGroupDark')} />
      <ThemeGroup kind="light" title={t('apGroupLight')} />

      <Group title={t('apGroupInterface')}>
        <label className="ap-row ap-row-click">
          <div className="ap-row-text">
            <div className="ap-row-label">{t('appearanceTransparency')}</div>
            <div className="ap-row-hint">{t('appearanceTransparencyHint')}</div>
          </div>
          <input
            type="checkbox"
            role="switch"
            className="set-switch"
            checked={appearance.transparency}
            onChange={(e) => updateAppearance({ transparency: e.target.checked })}
          />
        </label>
        <label className="ap-row ap-row-click">
          <div className="ap-row-text">
            <div className="ap-row-label">{t('appearanceConfirmDelete')}</div>
            <div className="ap-row-hint">{t('appearanceConfirmDeleteHint')}</div>
          </div>
          <input
            type="checkbox"
            role="switch"
            className="set-switch"
            checked={appearance.confirmDelete}
            onChange={(e) => updateAppearance({ confirmDelete: e.target.checked })}
          />
        </label>
      </Group>

      <div className="ap-footer">
        <span className="ap-footnote">{t('appearanceScope')}</span>
        <button type="button" className="set-btn" onClick={() => resetAppearance()}>
          {t('appearanceReset')}
        </button>
      </div>
    </>
  )
}

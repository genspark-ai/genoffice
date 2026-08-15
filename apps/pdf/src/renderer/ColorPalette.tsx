import type { CSSProperties, ReactElement } from 'react'

export interface ColorPreset {
  value: string
  label?: string
}

export function ColorPalette({
  value,
  presets,
  moreColorsLabel,
  onChange,
  columns = 5,
}: {
  value: string
  presets: readonly ColorPreset[]
  moreColorsLabel: string
  onChange: (value: string, source: 'preset' | 'custom') => void
  columns?: number
}): ReactElement {
  const selected = value.toUpperCase()
  const pickerValue = /^#[0-9A-F]{6}$/.test(selected) ? selected : '#000000'

  return (
    <div className="pdf-color-picker" style={{ '--pdf-color-columns': columns } as CSSProperties}>
      {presets.map((preset) => (
        <button
          key={preset.value}
          type="button"
          className={`pdf-color-swatch${selected === preset.value.toUpperCase() ? ' active' : ''}`}
          style={{ background: preset.value }}
          aria-label={preset.label ?? preset.value}
          onClick={() => onChange(preset.value, 'preset')}
        />
      ))}
      <label className="pdf-color-more">
        <input
          type="color"
          value={pickerValue}
          onChange={(event) => onChange(event.target.value, 'custom')}
        />
        <span>{moreColorsLabel}</span>
      </label>
    </div>
  )
}

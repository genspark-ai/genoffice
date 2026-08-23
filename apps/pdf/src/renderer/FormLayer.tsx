import type { ReactElement } from 'react'
import { Dropdown } from '@genoffice/ui'
import { pdfRectToCss } from './annotations'
import type { PageGeom } from './annotations'
import type { FormValueInput } from '../shared/ipc'
import type { FormWidget } from './form-catalog'
export {
  widgetsFromAnnotations,
  type FormWidget as Widget,
  type RawFormAnnotation,
} from './form-catalog'

/** AcroForm overlay: text fields/checkboxes absolutely positioned on the page; App aggregates values, ⌘S writes back */
export function FormLayer({
  widgets,
  geom,
  scale,
  readOnly = false,
  edits,
  activeWidgetId,
  signedWidgetIds,
  signatureLabel = 'Sign',
  registerControl,
  onFocus,
  onSignature,
  onEdit,
}: {
  widgets: FormWidget[]
  geom: PageGeom
  scale: number
  readOnly?: boolean
  edits: ReadonlyMap<string, FormValueInput>
  activeWidgetId?: string | null
  signedWidgetIds?: ReadonlySet<string>
  signatureLabel?: string
  registerControl?: (id: string, element: HTMLElement | null) => void
  onFocus?: (widget: FormWidget) => void
  onSignature?: (widget: FormWidget) => void
  onEdit: (value: FormValueInput) => void
}): ReactElement | null {
  if (widgets.length === 0) return null

  return (
    <div className="pdf-form-layer">
      {widgets.map((w) => {
        const style = pdfRectToCss(geom, w.rect, scale)
        const edit = edits.get(w.fieldName)
        const active = activeWidgetId === w.id ? ' is-active' : ''
        if (w.kind === 'signature') {
          const signed = w.signed || signedWidgetIds?.has(w.id)
          return (
            <button
              key={w.id}
              ref={(element) => registerControl?.(w.id, element)}
              type="button"
              className={`pdf-form-signature${signed ? ' is-signed' : ''}${active}`}
              style={style}
              disabled={readOnly || w.readOnly}
              aria-disabled={signed || undefined}
              aria-required={w.required}
              aria-label={signatureLabel}
              title={w.fieldName}
              onFocus={() => onFocus?.(w)}
              onClick={() => {
                if (!signed) onSignature?.(w)
              }}
            >
              {signed ? null : signatureLabel}
            </button>
          )
        }
        if (w.kind === 'checkbox') {
          const checked = edit ? !!edit.checked : w.checked
          return (
            <input
              key={w.id}
              ref={(element) => registerControl?.(w.id, element)}
              type="checkbox"
              className={`pdf-form-checkbox${w.required && !checked ? ' is-required-empty' : ''}${active}`}
              style={style}
              disabled={readOnly || w.readOnly}
              checked={checked}
              required={w.required}
              aria-required={w.required}
              aria-invalid={w.required && !checked}
              title={w.fieldName}
              onFocus={() => onFocus?.(w)}
              onChange={(e) =>
                onEdit({ name: w.fieldName, kind: 'checkbox', checked: e.target.checked })
              }
            />
          )
        }
        const value = edit ? (edit.value ?? '') : w.value
        if (w.kind === 'radio') {
          const missing = w.required && !value
          return (
            <input
              key={w.id}
              ref={(element) => registerControl?.(w.id, element)}
              type="radio"
              className={`pdf-form-radio${missing ? ' is-required-empty' : ''}${active}`}
              style={style}
              disabled={readOnly || w.readOnly}
              checked={value === w.buttonValue}
              name={w.fieldName}
              required={w.required}
              aria-required={w.required}
              aria-invalid={missing}
              title={w.fieldName}
              onFocus={() => onFocus?.(w)}
              onChange={() => onEdit({ name: w.fieldName, kind: 'radio', value: w.buttonValue })}
            />
          )
        }
        if (w.kind === 'choice') {
          const known = w.options.some((o) => o.exportValue === value)
          const missing = w.required && !value
          // Unknown current value: keep a disabled entry showing the raw export
          // value (the native select's hidden empty <option> equivalent) so the
          // list has no blank clickable row re-committing a bogus value.
          const options = [
            ...(known ? [] : [{ value, label: value, disabled: true }]),
            ...w.options.map((o) => ({ value: o.exportValue, label: o.displayValue })),
          ]
          return (
            <span
              key={w.id}
              // Focus (⇥ field navigation) must land on the inner trigger button
              ref={(element) =>
                registerControl?.(w.id, element?.querySelector<HTMLElement>('.gs-dd-btn') ?? null)
              }
              className={`pdf-form-select${missing ? ' is-required-empty' : ''}${active}`}
              style={{ ...style, fontSize: Math.max(9, Math.min(14, style.height * 0.55)) }}
              title={w.fieldName}
              onFocus={() => onFocus?.(w)}
            >
              <Dropdown
                value={value}
                options={options}
                disabled={readOnly || w.readOnly}
                ariaLabel={w.fieldName}
                ariaRequired={w.required || undefined}
                ariaInvalid={missing || undefined}
                onPick={(v) => onEdit({ name: w.fieldName, kind: 'choice', value: v })}
              />
            </span>
          )
        }
        const fontSize = Math.max(9, Math.min(14, style.height * 0.55))
        const missing = w.required && !value
        const className = `pdf-form-input${missing ? ' is-required-empty' : ''}${w.comb ? ' is-comb' : ''}${active}`
        if (w.multiLine) {
          return (
            <textarea
              key={w.id}
              ref={(element) => registerControl?.(w.id, element)}
              className={className}
              style={{ ...style, fontSize, textAlign: w.textAlignment }}
              disabled={readOnly || w.readOnly}
              value={value}
              maxLength={w.maxLen ?? undefined}
              required={w.required}
              aria-required={w.required}
              aria-invalid={missing}
              title={w.fieldName}
              onFocus={() => onFocus?.(w)}
              onChange={(e) => onEdit({ name: w.fieldName, kind: 'text', value: e.target.value })}
            />
          )
        }
        return (
          <input
            key={w.id}
            ref={(element) => registerControl?.(w.id, element)}
            type={w.password ? 'password' : 'text'}
            className={className}
            style={{ ...style, fontSize, textAlign: w.textAlignment }}
            disabled={readOnly || w.readOnly}
            value={value}
            maxLength={w.maxLen ?? undefined}
            required={w.required}
            aria-required={w.required}
            aria-invalid={missing}
            title={w.fieldName}
            onFocus={() => onFocus?.(w)}
            onChange={(e) => onEdit({ name: w.fieldName, kind: 'text', value: e.target.value })}
          />
        )
      })}
    </div>
  )
}

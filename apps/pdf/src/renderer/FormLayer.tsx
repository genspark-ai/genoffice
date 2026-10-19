import { useEffect, useRef, useState, type CSSProperties, type ReactElement } from 'react'
import { Dropdown } from '@genoffice/ui'
import { pdfRectToCss } from './annotations'
import type { PageGeom } from './annotations'
import type { FormValueInput } from '../shared/ipc'
import type { FormWidget } from './form-catalog'
import {
  formatDate,
  isDateOnlyFormat,
  isValidDateInput,
  normalizeDateInput,
  readDateParts,
  toIsoDate,
} from './date-format'
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
  dateInvalidLabel = (format: string) => format,
  datePickLabel = 'Pick a date',
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
  dateInvalidLabel?: (format: string) => string
  datePickLabel?: string
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
        if (w.dateFormat) {
          return (
            <DateInput
              key={w.id}
              widget={w}
              format={w.dateFormat}
              value={value}
              className={className}
              style={{ ...style, fontSize, textAlign: w.textAlignment }}
              disabled={readOnly || w.readOnly}
              missing={missing}
              invalidLabel={dateInvalidLabel(w.dateFormat)}
              pickLabel={datePickLabel}
              registerControl={registerControl}
              onFocus={onFocus}
              onEdit={onEdit}
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

/**
 * Acrobat-style date field. Complete dates in the field's format commit as typed; on leaving
 * the box other recognisable layouts are rewritten into the format, anything else is reverted
 * with the expected format shown under the box. A calendar button opens the native picker.
 */
function DateInput({
  widget,
  format,
  value,
  className,
  style,
  disabled,
  missing,
  invalidLabel,
  pickLabel,
  registerControl,
  onFocus,
  onEdit,
}: {
  widget: FormWidget
  format: string
  value: string
  className: string
  style: CSSProperties & { left: number; top: number; width: number; height: number }
  disabled: boolean
  missing: boolean
  invalidLabel: string
  pickLabel: string
  registerControl?: (id: string, element: HTMLElement | null) => void
  onFocus?: (widget: FormWidget) => void
  onEdit: (value: FormValueInput) => void
}): ReactElement {
  const [draft, setDraft] = useState<string | null>(null)
  const [rejected, setRejected] = useState(false)
  const pickerRef = useRef<HTMLInputElement>(null)
  useEffect(() => {
    if (!rejected) return
    const timer = window.setTimeout(() => setRejected(false), 3000)
    return () => window.clearTimeout(timer)
  }, [rejected])
  const commit = (next: string) => {
    if (next !== value) onEdit({ name: widget.fieldName, kind: 'text', value: next })
  }
  const { left, top, width, height } = style
  const picker = !disabled && height >= 10 && isDateOnlyFormat(format)
  const button = Math.min(height, 28)
  return (
    <>
      <input
        ref={(element) => registerControl?.(widget.id, element)}
        type="text"
        className={`${className}${rejected ? ' is-invalid' : ''}`}
        style={picker ? { ...style, paddingRight: button } : style}
        disabled={disabled}
        value={draft ?? value}
        placeholder={format}
        required={widget.required}
        aria-required={widget.required}
        aria-invalid={missing || rejected}
        title={widget.fieldName}
        onFocus={() => {
          setRejected(false)
          onFocus?.(widget)
        }}
        onChange={(e) => {
          const next = e.target.value
          setDraft(next)
          if (next.trim() && isValidDateInput(next, format)) commit(next.trim())
        }}
        onKeyDown={(e) => {
          if (e.key === 'Enter') e.currentTarget.blur()
        }}
        onBlur={() => {
          if (draft !== null) {
            const normalized = normalizeDateInput(draft, format)
            if (normalized === null) setRejected(true)
            else commit(normalized)
          }
          setDraft(null)
        }}
      />
      {picker && (
        <>
          <button
            type="button"
            className="pdf-form-date-pick"
            style={{
              left: left + width - button,
              top: top + (height - button) / 2,
              width: button,
              height: button,
            }}
            tabIndex={-1}
            title={pickLabel}
            aria-label={pickLabel}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              const native = pickerRef.current
              if (!native) return
              const parts = readDateParts(value, format)
              native.value = parts ? toIsoDate(parts) : ''
              onFocus?.(widget)
              try {
                native.showPicker()
              } catch {
                native.click()
              }
            }}
          >
            <svg
              viewBox="0 0 16 16"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.4"
              aria-hidden
            >
              <rect x="2" y="3" width="12" height="11" rx="1.5" />
              <path d="M2 6.5h12M5.5 1.5v3M10.5 1.5v3" />
            </svg>
          </button>
          <input
            ref={pickerRef}
            type="date"
            className="pdf-form-date-native"
            style={{ left: left + width - button, top, width: button, height }}
            tabIndex={-1}
            aria-hidden
            onChange={(e) => {
              const [y, m, d] = e.target.value.split('-').map(Number)
              if (!y || !m || !d) return
              setDraft(null)
              setRejected(false)
              commit(formatDate({ year: y, month: m, day: d }, format))
            }}
          />
        </>
      )}
      {rejected && (
        <span role="alert" className="pdf-form-date-error" style={{ left, top: top + height + 2 }}>
          {invalidLabel}
        </span>
      )}
    </>
  )
}

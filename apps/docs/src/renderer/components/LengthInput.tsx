import { useEffect, useRef } from 'react'
import type { FocusEvent, FormEvent, KeyboardEvent } from 'react'
import { useMeasurement } from '../use-measurement'
import { UNIT_STEP } from '../units'

/**
 * Word-style length field: shows the value in the preferred unit with its
 * suffix (`2.54 cm`, `1"`), accepts any unit on input (`36pt`, `1 in`) and
 * commits twips on Enter / blur; arrow keys step like Word's spinner. The
 * field is uncontrolled so a half-typed `-` or `2.` survives re-renders.
 */
export function LengthInput({
  value,
  onCommit,
  min = 0,
  max,
  allowEmpty = false,
  live = false,
  placeholder,
  disabled,
  ariaLabel,
  className,
  autoFocus,
  onEnter,
  onFocus,
  onBlur,
}: {
  /** twips; null renders empty (e.g. "Auto") */
  value: number | null
  onCommit: (twips: number | null) => void
  /** clamp bounds in twips */
  min?: number
  max?: number
  /** an emptied field commits null instead of reverting */
  allowEmpty?: boolean
  /** dialogs: commit every readable keystroke so OK never reads a stale value
   *  (the text is only re-formatted on blur); ribbon fields keep Enter/blur */
  live?: boolean
  placeholder?: string
  disabled?: boolean
  ariaLabel?: string
  className?: string
  autoFocus?: boolean
  /** after the Enter commit (the parent's state is not yet flushed here) */
  onEnter?: () => void
  onFocus?: (e: FocusEvent<HTMLInputElement>) => void
  onBlur?: (e: FocusEvent<HTMLInputElement>) => void
}) {
  const m = useMeasurement()
  const ref = useRef<HTMLInputElement>(null)
  // Enter commits and then blurs: the blur must not commit the same value again
  const committedRef = useRef(value)
  const text = value === null ? '' : m.format(value)

  useEffect(() => {
    committedRef.current = value
  }, [value])
  useEffect(() => {
    const el = ref.current
    if (el && document.activeElement !== el) el.value = text
  }, [text])

  const emit = (next: number | null) => {
    if (next === committedRef.current) return
    committedRef.current = next
    onCommit(next)
  }

  const clamp = (twips: number) => {
    let v = Math.max(min, twips)
    if (max !== undefined) v = Math.min(max, v)
    return v
  }

  // the rounded display of the stored or just-committed value: re-parsing it
  // (Enter, then blur before the parent re-rendered) must not drift the twips
  const isShownText = (raw: string) =>
    raw === text || raw === (committedRef.current === null ? '' : m.format(committedRef.current))

  const commit = (el: HTMLInputElement): void => {
    const raw = el.value.trim()
    if (isShownText(raw)) return
    if (raw === '' && allowEmpty) {
      emit(null)
      return
    }
    const parsed = raw === '' ? null : m.parse(raw)
    if (parsed === null) {
      el.value = text
      return
    }
    const next = clamp(parsed)
    el.value = m.format(next)
    emit(next)
  }

  const onInput = (e: FormEvent<HTMLInputElement>) => {
    if (!live) return
    const raw = e.currentTarget.value.trim()
    if (isShownText(raw)) return
    if (raw === '' && allowEmpty) emit(null)
    else if (raw !== '') {
      const parsed = m.parse(raw)
      if (parsed !== null) emit(clamp(parsed))
    }
  }

  const onKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
    const el = e.currentTarget
    if (e.key === 'Enter') {
      e.preventDefault()
      commit(el)
      onEnter?.()
      return
    }
    if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
      e.preventDefault()
      const typed = el.value.trim()
      const base = isShownText(typed)
        ? (committedRef.current ?? 0)
        : (m.parse(typed) ?? committedRef.current ?? 0)
      const step = m.fromUnit(UNIT_STEP[m.unit]) * (e.key === 'ArrowUp' ? 1 : -1)
      // step lands on the unit grid, not on an off-grid stored value
      const stepped = Math.round((base + step) / Math.abs(step)) * Math.abs(step)
      const next = clamp(stepped)
      el.value = m.format(next)
      emit(next)
    }
  }

  return (
    <input
      ref={ref}
      type="text"
      inputMode="decimal"
      className={className ? `len-input ${className}` : 'len-input'}
      defaultValue={text}
      placeholder={placeholder}
      disabled={disabled}
      aria-label={ariaLabel}
      autoFocus={autoFocus}
      data-unit={m.unit}
      onInput={onInput}
      onKeyDown={onKeyDown}
      onFocus={onFocus}
      onBlur={(e) => {
        commit(e.currentTarget)
        onBlur?.(e)
      }}
    />
  )
}

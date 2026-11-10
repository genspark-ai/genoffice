import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import type { KeyboardEvent, RefObject } from 'react'
import {
  MAX_LABEL_LENGTH,
  placeholderSource,
  sanitizeLabel,
} from '@genoffice/agent-core/redact-core'

/**
 * The label field of a "hide from the model" dialog.
 *
 * Five editors each grew this dialog, and the behaviour came out identical
 * five times: the label is sanitised as it is typed, the `{{placeholder}}` the
 * model will see is shown live so the reader knows what they are about to
 * expose before committing, the input is capped and selected on open, Enter
 * submits, Escape cancels, and the primary button is dead while the label is
 * empty.
 *
 * That behaviour is where a bug would live — five copies means five places to
 * forget that an empty label must not submit, or that the preview has to be the
 * sanitised text rather than what was typed. The *markup* is not shared: each
 * editor wears its own modal chrome and its own class names, which is a
 * deliberate choice recorded in each dialog, so that part stays where it is.
 */

export interface RedactLabelField {
  value: string
  setValue: (next: string) => void
  /** the label as it will be stored, and as the model will see it named */
  label: string
  /** `{{label}}`, live, or the private fallback when the field is empty */
  marker: string
  /** true while there is nothing worth submitting */
  empty: boolean
  inputRef: RefObject<HTMLInputElement | null>
  onKeyDown: (event: KeyboardEvent<HTMLInputElement>) => void
  /** the input's `maxLength`, so the cap is stated in one place too */
  maxLength: number
}

export function useRedactLabelField(
  seed: string,
  onSubmit: (label: string) => void,
  onCancel: () => void,
): RedactLabelField {
  const [value, setValue] = useState(seed)
  const inputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    inputRef.current?.focus()
    inputRef.current?.select()
  }, [])

  const label = useMemo(() => sanitizeLabel(value), [value])

  const onKeyDown = useCallback(
    (event: KeyboardEvent<HTMLInputElement>) => {
      if (event.key === 'Enter' && sanitizeLabel(value) !== '') {
        // the sanitised label, not the typed one: what the reader typed may
        // carry characters the carrier cannot hold
        event.preventDefault()
        onSubmit(sanitizeLabel(value))
      } else if (event.key === 'Escape') {
        event.preventDefault()
        onCancel()
      }
    },
    [onCancel, onSubmit, value],
  )

  return {
    value,
    setValue,
    label,
    marker: placeholderSource(label),
    empty: label === '',
    inputRef,
    onKeyDown,
    maxLength: MAX_LABEL_LENGTH,
  }
}

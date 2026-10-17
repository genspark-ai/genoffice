import { useEffect, useMemo, useRef } from 'react'
import type { FormEvent } from 'react'
import { useI18n } from '../i18n/locale'
import { DebouncedSearch } from './list-perf'
import './list.css'

/**
 * Mailbox search box (Lane R owns; H6 debounce P-4). Submit runs `mail:search-local`
 * over the IPC fakes, then the caller applies the `matchesQuery`
 * post-filter locally (`applyLocalSearch`); the real sync engine is a
 * later phase. Server expansion stays an explicit button in the list
 * footer (`mail.list.moreOnServer`), never implicit.
 *
 * Debounce (opt-in): pass `debounceMs` (P-4 pins 150-250) with
 * `onDebouncedQuery` for search-as-you-type. Every keystroke still
 * calls `onChange` immediately (the input stays responsive) and also
 * restarts the debounce timer; a superseded pending query fires
 * `onQuerySuperseded` so the caller can cancel the in-flight store
 * query (H4 owns the main-side cancellation contract). With
 * `debounceMs` unset the box keeps the legacy submit-only behavior.
 */

export interface SearchBoxProps {
  value: string
  searching: boolean
  onChange(value: string): void
  onSubmit(): void
  onClear(): void
  /** Opt-in debounce delay in ms (P-4: 150-250); unset keeps legacy behavior. */
  debounceMs?: number
  /** Fires with the settled query + generation after the debounce delay. */
  onDebouncedQuery?(query: string, generation: number): void
  /** Fires when a keystroke cancels a pending debounced query. */
  onQuerySuperseded?(): void
}

export function SearchBox({
  value,
  searching,
  onChange,
  onSubmit,
  onClear,
  debounceMs = 0,
  onDebouncedQuery,
  onQuerySuperseded,
}: SearchBoxProps) {
  const { t } = useI18n()
  // Latest callbacks without rebuilding the debouncer each render.
  const fireRef = useRef(onDebouncedQuery)
  fireRef.current = onDebouncedQuery
  const supersededRef = useRef(onQuerySuperseded)
  supersededRef.current = onQuerySuperseded
  const debouncer = useMemo(
    () =>
      debounceMs > 0
        ? new DebouncedSearch(
            debounceMs,
            (query, generation) => fireRef.current?.(query, generation),
            () => supersededRef.current?.(),
          )
        : null,
    [debounceMs],
  )
  useEffect(() => {
    if (debouncer) debouncer.push(value)
  }, [value, debouncer])
  useEffect(() => () => debouncer?.cancel(), [debouncer])

  const submit = (event: FormEvent): void => {
    event.preventDefault()
    onSubmit()
  }

  return (
    <form className="mail-search" role="search" onSubmit={submit}>
      <label className="mail-search-label" htmlFor="mail-search-input">
        {t('mail.search.box')}
      </label>
      <input
        id="mail-search-input"
        className="mail-search-input"
        type="search"
        autoComplete="off"
        placeholder={t('mail.list.searchPlaceholder')}
        value={value}
        disabled={searching}
        onChange={(event) => onChange(event.target.value)}
      />
      {value && (
        <button
          type="button"
          className="mail-search-clear"
          aria-label={t('mail.search.box')}
          onClick={onClear}
        >
          x
        </button>
      )}
      <button type="submit" className="mail-search-go" disabled={searching}>
        {t('mail.search.box')}
      </button>
    </form>
  )
}

import { useEffect, useRef, useState, type KeyboardEvent, type MouseEvent } from 'react'
import type { ZoteroLocalItem } from '../../shared/ipc'
import { useI18n, type StringKey } from '../i18n/locale'
import { IS_MAC } from '../keybindings'
import { ZoteroItemText, useZoteroItems } from './zotero-items'

export interface ZoteroQuickCiteProps {
  /** cites the items at the caret; resolves to the message to show when it could not */
  onCite(keys: string[]): Promise<StringKey | null>
  onClose(): void
}

/**
 * Quick Cite: a search bar over the document, like Zotero's own Quick Format
 * bar, without leaving GenOffice. Type part of a title, an author or a year;
 * Enter cites the highlighted item, and ⌘-click (Ctrl-click) or ⌘-Enter picks
 * several for one citation.
 */
export function ZoteroQuickCite({ onCite, onClose }: ZoteroQuickCiteProps) {
  const { t } = useI18n()
  const [query, setQuery] = useState('')
  const [active, setActive] = useState(0)
  const [picked, setPicked] = useState<ZoteroLocalItem[]>([])
  const [citing, setCiting] = useState(false)
  const [citeError, setCiteError] = useState<StringKey | null>(null)
  const inputRef = useRef<HTMLInputElement>(null)
  const rootRef = useRef<HTMLDivElement>(null)
  const trimmed = query.trim()
  const {
    items,
    loading,
    error,
    key: answered,
  } = useZoteroItems(trimmed, () => window.desktop.zoteroSearch(trimmed))

  // back to the search box after a citation that could not go in (the box was disabled meanwhile)
  useEffect(() => {
    if (!citing) inputRef.current?.focus()
  }, [citing])
  useEffect(() => setActive(0), [items])
  // a click elsewhere puts the bar away, as Escape does
  useEffect(() => {
    const away = (event: globalThis.MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) onClose()
    }
    document.addEventListener('mousedown', away)
    return () => document.removeEventListener('mousedown', away)
  }, [onClose])

  const togglePick = (item: ZoteroLocalItem) =>
    setPicked((prev) =>
      prev.some((p) => p.key === item.key)
        ? prev.filter((p) => p.key !== item.key)
        : [...prev, item],
    )

  const cite = async (also?: ZoteroLocalItem) => {
    const keys = picked.map((item) => item.key)
    if (also && !keys.includes(also.key)) keys.push(also.key)
    if (keys.length === 0 || citing) return
    setCiting(true)
    setCiteError(null)
    const failure = await onCite(keys)
    setCiting(false)
    if (failure) setCiteError(failure)
    else onClose()
  }

  const onKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'ArrowDown') {
      event.preventDefault()
      setActive((i) => Math.min(items.length - 1, i + 1))
    } else if (event.key === 'ArrowUp') {
      event.preventDefault()
      setActive((i) => Math.max(0, i - 1))
    } else if (event.key === 'Enter') {
      event.preventDefault()
      // the list still shows the previous search: Enter would cite the wrong item
      if (answered !== trimmed) return
      const item = items[active]
      if ((event.metaKey || event.ctrlKey) && item) togglePick(item)
      else void cite(item)
    } else if (event.key === 'Backspace' && query === '' && picked.length > 0) {
      setPicked((prev) => prev.slice(0, -1))
    }
  }

  const onRowClick = (event: MouseEvent, item: ZoteroLocalItem) => {
    if (event.metaKey || event.ctrlKey) togglePick(item)
    else void cite(item)
    inputRef.current?.focus()
  }

  const message: StringKey | null =
    citeError ?? error ?? (trimmed && !loading && items.length === 0 ? 'zoteroNoResults' : null)

  return (
    <div
      className="zotero-quick-cite"
      ref={rootRef}
      role="dialog"
      aria-label={t('zoteroQuickCite')}
      onKeyDown={(event) => {
        if (event.key !== 'Escape') return
        event.preventDefault()
        onClose()
      }}
    >
      <div className="zotero-quick-cite-bar">
        {picked.map((item) => (
          <button
            key={item.key}
            type="button"
            className="zotero-quick-cite-chip"
            data-tip={item.title}
            // command-exempt: takes an item back out of the citation being built
            onClick={() => togglePick(item)}
          >
            {item.creators || item.title} {item.year}
          </button>
        ))}
        <input
          ref={inputRef}
          value={query}
          disabled={citing}
          placeholder={t('zoteroSearchPlaceholder')}
          aria-label={t('zoteroQuickCite')}
          data-testid="zotero-quick-cite-input"
          onChange={(event) => {
            setQuery(event.target.value)
            setCiteError(null)
          }}
          onKeyDown={onKeyDown}
        />
      </div>
      {message ? (
        <div className="zotero-quick-cite-message">{t(message)}</div>
      ) : (
        items.length > 0 && (
          <div className="zotero-quick-cite-list" role="listbox">
            {items.map((item, index) => {
              const isPicked = picked.some((p) => p.key === item.key)
              return (
                <button
                  key={item.key}
                  type="button"
                  role="option"
                  aria-selected={index === active}
                  className={`selection-pane-row${index === active ? ' active' : ''}${isPicked ? ' picked' : ''}`}
                  onMouseEnter={() => setActive(index)}
                  // command-exempt: picks the items of the citation, like typing in the bar
                  onClick={(event) => onRowClick(event, item)}
                >
                  <ZoteroItemText item={item} />
                </button>
              )
            })}
          </div>
        )
      )}
      <div className="zotero-quick-cite-hint">
        {t('zoteroQuickCiteHint', { mod: IS_MAC ? '⌘' : 'Ctrl' })}
      </div>
    </div>
  )
}

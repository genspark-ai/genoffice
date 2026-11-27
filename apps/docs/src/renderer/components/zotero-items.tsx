import { useEffect, useState } from 'react'
import type { ZoteroLocalItem, ZoteroLocalResult } from '../../shared/ipc'
import type { StringKey } from '../i18n/locale'
import { zoteroLocalErrorKey } from '../zotero/quick-cite'

/** wait after the last keystroke before asking Zotero */
export const ZOTERO_SEARCH_DEBOUNCE_MS = 200

export interface ZoteroItemsState {
  items: ZoteroLocalItem[]
  loading: boolean
  /** the message to show instead of the list */
  error: StringKey | null
  /** the request these items answer: while a newer one loads they are the previous answer */
  key: string | null
}

/**
 * The items a lookup returns, kept to the latest request: a slower answer to an
 * older query never replaces a newer one. `key` names the request; the lookup
 * runs again whenever it changes.
 */
export function useZoteroItems(
  key: string,
  load: () => Promise<ZoteroLocalResult<ZoteroLocalItem[]>>,
  delay = ZOTERO_SEARCH_DEBOUNCE_MS,
): ZoteroItemsState {
  const [state, setState] = useState<ZoteroItemsState>({
    items: [],
    loading: false,
    error: null,
    key: null,
  })
  useEffect(() => {
    let live = true
    setState((prev) => ({ ...prev, loading: true }))
    const timer = setTimeout(() => {
      load().then(
        (result) => {
          if (!live) return
          if (result.ok) setState({ items: result.value, loading: false, error: null, key })
          else {
            console.error('Zotero local API failed', result.error)
            setState({
              items: [],
              loading: false,
              error: zoteroLocalErrorKey(result.errorCode),
              key,
            })
          }
        },
        (error: unknown) => {
          if (!live) return
          console.error('Zotero local API failed', error)
          setState({ items: [], loading: false, error: 'zoteroOperationError', key })
        },
      )
    }, delay)
    return () => {
      live = false
      clearTimeout(timer)
    }
    // `load` is rebuilt every render; `key` is what names the request
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key, delay])
  return state
}

/** "Smith and Lee · 2020", the line under an item's title */
export function zoteroItemByline(item: ZoteroLocalItem): string {
  return [item.creators, item.year].filter(Boolean).join(' · ')
}

/** one item of a list: its title over its creators and year */
export function ZoteroItemText({ item }: { item: ZoteroLocalItem }) {
  return (
    <>
      <span className="zotero-item-title">{item.title || item.key}</span>
      <span className="selection-pane-text">{zoteroItemByline(item)}</span>
    </>
  )
}

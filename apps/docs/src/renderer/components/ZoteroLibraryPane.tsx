import { useEffect, useMemo, useState, type MouseEvent } from 'react'
import type { ZoteroLocalCollection, ZoteroLocalItem } from '../../shared/ipc'
import { useI18n, type StringKey } from '../i18n/locale'
import { zoteroLocalErrorKey } from '../zotero/quick-cite'
import { ZoteroItemText, useZoteroItems } from './zotero-items'

/** a list box collapses ordinary spaces, so nesting is shown with no-break ones */
const INDENT = '\u00a0\u00a0\u00a0'

/** collections in tree order, each with its depth, parents before their children */
export function collectionTree(
  collections: ZoteroLocalCollection[],
): Array<{ collection: ZoteroLocalCollection; depth: number }> {
  const keys = new Set(collections.map((c) => c.key))
  const children = new Map<string | null, ZoteroLocalCollection[]>()
  for (const collection of collections) {
    // a parent outside the list (a trashed one) puts the collection at the top
    const parent = collection.parent && keys.has(collection.parent) ? collection.parent : null
    children.set(parent, [...(children.get(parent) ?? []), collection])
  }
  const out: Array<{ collection: ZoteroLocalCollection; depth: number }> = []
  const seen = new Set<string>()
  const walk = (parent: string | null, depth: number) => {
    for (const collection of children.get(parent) ?? []) {
      if (seen.has(collection.key)) continue
      seen.add(collection.key)
      out.push({ collection, depth })
      walk(collection.key, depth + 1)
    }
  }
  walk(null, 0)
  return out
}

export interface ZoteroLibraryPaneProps {
  /** cites the items at the caret; resolves to the message to show when it could not */
  onCite(keys: string[]): Promise<StringKey | null>
  onClose(): void
}

/**
 * The Zotero library beside the document: pick a collection, filter it, and
 * click an item to cite it at the caret. ⌘-click (Ctrl-click) picks several
 * for one citation.
 */
export function ZoteroLibraryPane({ onCite, onClose }: ZoteroLibraryPaneProps) {
  const { t } = useI18n()
  const [collections, setCollections] = useState<ZoteroLocalCollection[]>([])
  const [collectionsError, setCollectionsError] = useState<StringKey | null>(null)
  const [collection, setCollection] = useState<string | null>(null)
  const [query, setQuery] = useState('')
  const [picked, setPicked] = useState<string[]>([])
  const [citing, setCiting] = useState(false)
  const [citeError, setCiteError] = useState<StringKey | null>(null)
  // the library is read again when the user comes back from Zotero, where items may have changed
  const [reload, setReload] = useState(0)

  useEffect(() => {
    const again = () => setReload((n) => n + 1)
    window.addEventListener('focus', again)
    return () => window.removeEventListener('focus', again)
  }, [])

  useEffect(() => {
    let live = true
    void window.desktop.zoteroCollections().then(
      (result) => {
        if (!live) return
        if (result.ok) {
          setCollections(result.value)
          setCollectionsError(null)
        } else {
          console.error('Zotero local API failed', result.error)
          setCollectionsError(zoteroLocalErrorKey(result.errorCode))
        }
      },
      (error: unknown) => {
        if (live) {
          console.error('Zotero local API failed', error)
          setCollectionsError('zoteroOperationError')
        }
      },
    )
    return () => {
      live = false
    }
  }, [reload])

  const trimmed = query.trim()
  const { items, loading, error } = useZoteroItems(
    `${collection ?? ''}\n${trimmed}\n${reload}`,
    () => window.desktop.zoteroCollectionItems(collection, trimmed),
  )
  const tree = useMemo(() => collectionTree(collections), [collections])
  // a collection deleted in Zotero falls back to the whole library
  useEffect(() => {
    if (collection && collections.length > 0 && !collections.some((c) => c.key === collection))
      setCollection(null)
  }, [collection, collections])

  const cite = async (keys: string[]) => {
    if (keys.length === 0 || citing) return
    setCiting(true)
    setCiteError(null)
    const failure = await onCite(keys)
    setCiting(false)
    if (failure) setCiteError(failure)
    else setPicked([])
  }

  const onRowClick = (event: MouseEvent, item: ZoteroLocalItem) => {
    if (event.metaKey || event.ctrlKey) {
      setPicked((prev) =>
        prev.includes(item.key) ? prev.filter((key) => key !== item.key) : [...prev, item.key],
      )
    } else {
      void cite(picked.includes(item.key) ? picked : [...picked, item.key])
    }
  }

  const message: StringKey | null =
    citeError ??
    collectionsError ??
    error ??
    (!loading && items.length === 0 ? 'zoteroNoResults' : null)

  return (
    <aside className="comments-pane zotero-library-pane" data-testid="zotero-library-pane">
      <div className="comments-pane-head">
        <span className="comments-pane-title">{t('zoteroLibrary')}</span>
        <button
          className="comments-pane-close"
          data-tip={t('appClose')}
          aria-label={t('appClose')}
          onClick={onClose}
        >
          ✕
        </button>
      </div>
      <div className="zotero-library-controls">
        <select
          value={collection ?? ''}
          aria-label={t('zoteroLibrary')}
          onChange={(event) => {
            setCollection(event.target.value || null)
            setPicked([])
          }}
        >
          <option value="">{t('zoteroMyLibrary')}</option>
          {tree.map(({ collection: c, depth }) => (
            <option key={c.key} value={c.key}>
              {`${INDENT.repeat(depth)}${c.name}`}
            </option>
          ))}
        </select>
        <input
          type="search"
          value={query}
          placeholder={t('zoteroSearchPlaceholder')}
          aria-label={t('zoteroSearchPlaceholder')}
          onChange={(event) => {
            setQuery(event.target.value)
            setCiteError(null)
          }}
        />
      </div>
      <div className="comments-pane-list" role="listbox" aria-label={t('zoteroLibrary')}>
        {message && <div className="comments-empty">{t(message)}</div>}
        {!message &&
          items.map((item) => {
            const isPicked = picked.includes(item.key)
            return (
              <button
                key={item.key}
                type="button"
                role="option"
                aria-selected={isPicked}
                disabled={citing}
                className={`selection-pane-row${isPicked ? ' active' : ''}`}
                data-tip={t('zoteroCiteTip')}
                // command-exempt: picks the items to cite, like the Quick Cite bar
                onClick={(event) => onRowClick(event, item)}
              >
                <ZoteroItemText item={item} />
              </button>
            )
          })}
      </div>
      {picked.length > 0 && (
        <div className="zotero-library-actions">
          <button
            type="button"
            className="btn-primary"
            disabled={citing}
            // command-exempt: cites the items picked in the pane
            onClick={() => void cite(picked)}
          >
            {t('zoteroCitePicked', { n: picked.length })}
          </button>
        </div>
      )}
    </aside>
  )
}

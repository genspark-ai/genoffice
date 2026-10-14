import { useEffect, useState } from 'react'
import type { ReactElement } from 'react'
import { useI18n } from '../i18n/locale'
import type { WikiBacklink } from '../../shared/ipc'

/**
 * The backlinks sidebar: notes whose links land on the open document. The
 * index is derived — nothing is stored in the vault — so the list is exactly
 * what an Obsidian-compatible vault's own backlinks show for the same files.
 * A row opens its source note; the header button re-walks when the vault
 * changed elsewhere.
 */
export function BacklinksPane({
  path,
  onOpen,
}: {
  /** absolute path of the open document; null while untitled */
  path: string | null
  onOpen: (path: string) => void
}): ReactElement {
  const { t } = useI18n()
  const [links, setLinks] = useState<WikiBacklink[] | null>(null)
  const [reload, setReload] = useState(0)

  useEffect(() => {
    let cancelled = false
    setLinks(null)
    // The browser test harness has no Electron channel; an absent API is an
    // empty list, not a failed pane.
    if (!path || typeof window.markdownApi?.wikiBacklinks !== 'function') {
      setLinks([])
      return
    }
    window.markdownApi
      .wikiBacklinks(path)
      .then((result) => {
        if (!cancelled) setLinks(result)
      })
      .catch(() => {
        if (!cancelled) setLinks([])
      })
    return () => {
      cancelled = true
    }
  }, [path, reload])

  const name = (source: string): string => {
    const file = source.replaceAll('\\', '/').split('/').pop() ?? source
    const dot = file.lastIndexOf('.')
    return dot > 0 ? file.slice(0, dot) : file
  }

  return (
    <aside className="md-backlinks" aria-label={t('backlinks')}>
      <div className="md-backlinks-head">
        <span className="md-backlinks-title">{t('backlinks')}</span>
        <button
          type="button"
          className="md-backlinks-refresh"
          aria-label={t('backlinksRefresh')}
          title={t('backlinksRefresh')}
          onClick={() => setReload((n) => n + 1)}
        >
          ⟳
        </button>
      </div>
      {links === null ? null : links.length === 0 ? (
        <div className="md-backlinks-empty">{t('backlinksEmpty')}</div>
      ) : (
        <nav className="md-backlinks-list">
          {links.map((link, index) => (
            <button
              key={`${link.source}-${link.target}-${index}`}
              type="button"
              className="md-backlinks-item"
              title={link.source}
              onClick={() => onOpen(link.source)}
            >
              <span className="md-backlinks-name">{name(link.source)}</span>
              <span className="md-backlinks-target">{link.target}</span>
            </button>
          ))}
        </nav>
      )}
    </aside>
  )
}

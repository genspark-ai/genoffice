import { useEffect, useState } from 'react'
import type { ReactElement } from 'react'
import { useI18n } from '../i18n/locale'

/**
 * The tags sidebar: every #tag in the vault with the notes carrying it, most
 * used first. The index is derived — nothing is stored in the vault — so the
 * list is exactly what the files' own text says. A tag expands to its notes;
 * a note row opens it.
 */
export function TagsPane({ onOpen }: { onOpen: (path: string) => void }): ReactElement {
  const { t } = useI18n()
  const [tags, setTags] = useState<{ tag: string; files: string[] }[] | null>(null)
  const [openTag, setOpenTag] = useState<string | null>(null)
  const [reload, setReload] = useState(0)

  useEffect(() => {
    let cancelled = false
    setTags(null)
    // The browser test harness has no Electron channel; an absent API is an
    // empty list, not a failed pane.
    if (typeof window.markdownApi?.wikiTags !== 'function') {
      setTags([])
      return
    }
    window.markdownApi
      .wikiTags()
      .then((result) => {
        if (!cancelled) setTags(result)
      })
      .catch(() => {
        if (!cancelled) setTags([])
      })
    return () => {
      cancelled = true
    }
  }, [reload])

  const name = (source: string): string => {
    const file = source.replaceAll('\\', '/').split('/').pop() ?? source
    const dot = file.lastIndexOf('.')
    return dot > 0 ? file.slice(0, dot) : file
  }

  return (
    <aside className="md-backlinks" aria-label={t('tags')}>
      <div className="md-backlinks-head">
        <span className="md-backlinks-title">{t('tags')}</span>
        <button
          type="button"
          className="md-backlinks-refresh"
          aria-label={t('backlinksRefresh')}
          title={t('backlinksRefresh')}
          onClick={() => setReload((v) => v + 1)}
        >
          ⟳
        </button>
      </div>
      {tags === null ? (
        <div className="md-backlinks-empty">{t('loading')}</div>
      ) : tags.length === 0 ? (
        <div className="md-backlinks-empty">{t('tagsEmpty')}</div>
      ) : (
        <ul className="md-backlinks-list">
          {tags.map(({ tag, files }) => (
            <li key={tag}>
              <button
                type="button"
                className="md-backlinks-item"
                onClick={() => setOpenTag((v) => (v === tag ? null : tag))}
              >
                <span className="md-backlinks-name">#{tag}</span>
                <span className="md-backlinks-target">{files.length}</span>
              </button>
              {openTag === tag && (
                <ul className="md-backlinks-list md-tags-notes">
                  {files.map((file) => (
                    <li key={file}>
                      <button
                        type="button"
                        className="md-backlinks-item"
                        title={file}
                        onClick={() => onOpen(file)}
                      >
                        <span className="md-backlinks-name">{name(file)}</span>
                      </button>
                    </li>
                  ))}
                </ul>
              )}
            </li>
          ))}
        </ul>
      )}
    </aside>
  )
}

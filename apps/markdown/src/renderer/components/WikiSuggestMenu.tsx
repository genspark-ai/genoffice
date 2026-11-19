import { forwardRef, useEffect, useImperativeHandle, useState } from 'react'
import type { WikiSuggestState } from '../editor/wikiSuggest'
import { useI18n } from '../i18n/locale'

export interface WikiSuggestMenuHandle {
  handleKey(event: KeyboardEvent): boolean
}

interface Props {
  state: WikiSuggestState | null
  onDismiss: () => void
}

const NO_ITEMS: WikiSuggestState['items'] = []

/**
 * The `[[` suggester's popup: vault notes, shallowest first, filtered live,
 * and — while the query names something no note has — a last row that creates
 * that note (Obsidian's "Create new note"), keyboard and click both fine.
 * After a `#` the rows are the target's headings (`##`-glyphed) or its
 * `^block` ids with a line of context.
 */
export const WikiSuggestMenu = forwardRef<WikiSuggestMenuHandle, Props>(function WikiSuggestMenu(
  { state, onDismiss },
  ref,
) {
  const { t } = useI18n()
  const [index, setIndex] = useState(0)

  const items = state?.items ?? NO_ITEMS
  const canCreate = (state?.createName ?? '') !== ''
  const createAt = items.length
  const count = items.length + (canCreate ? 1 : 0)

  useEffect(() => {
    setIndex(0)
  }, [state?.items])

  useImperativeHandle(
    ref,
    () => ({
      handleKey(event: KeyboardEvent): boolean {
        if (!state) return false
        if (event.key === 'ArrowDown') {
          setIndex((i) => (count === 0 ? 0 : (i + 1) % count))
          return true
        }
        if (event.key === 'ArrowUp') {
          setIndex((i) => (count === 0 ? 0 : (i - 1 + count) % count))
          return true
        }
        if (event.key === 'Enter') {
          const item = items[index]
          if (item) state.command(item)
          else if (canCreate) state.create()
          return true
        }
        if (event.key === 'Escape') {
          onDismiss()
          return true
        }
        return false
      },
    }),
    [state, index, count, items, canCreate, onDismiss],
  )

  if (!state || !state.clientRect) return null
  const rect = state.clientRect

  return (
    <div
      className="slash-menu wiki-suggest"
      style={{ position: 'fixed', top: rect.bottom + 4, left: rect.left }}
      onMouseDown={(e) => e.preventDefault()}
    >
      {count === 0 ? (
        <div className="slash-empty">{t('slashNoResults')}</div>
      ) : (
        <>
          {items.map((item, i) => {
            const row =
              item.kind === 'note'
                ? { name: item.entry.name, path: item.entry.relPath, title: item.entry.relPath }
                : item.kind === 'heading'
                  ? { name: item.text, path: '#'.repeat(item.level), title: item.text }
                  : { name: `^${item.id}`, path: item.text, title: item.text }
            return (
              <button
                key={`${item.kind}:${i}`}
                type="button"
                className={`slash-item wiki-suggest-item${i === index ? ' active' : ''}`}
                title={row.title}
                onMouseEnter={() => setIndex(i)}
                onClick={() => state.command(item)}
              >
                <span className="wiki-suggest-name">{row.name}</span>
                <span className="wiki-suggest-path">{row.path}</span>
              </button>
            )
          })}
          {canCreate && (
            <button
              type="button"
              className={`slash-item wiki-suggest-item${createAt === index ? ' active' : ''}`}
              title={state.createName}
              onMouseEnter={() => setIndex(createAt)}
              onClick={() => state.create()}
            >
              <span className="wiki-suggest-name">{t('suggestCreateNote')}</span>
              <span className="wiki-suggest-path">{state.createName}</span>
            </button>
          )}
        </>
      )}
    </div>
  )
})

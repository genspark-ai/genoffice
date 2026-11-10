import type { ReactElement } from 'react'
import { ancestorsOf, type ParseMap } from '../document/parse-map'
import { useI18n } from '../i18n/locale'

export type NodeState = 'static' | 'dynamic' | 'dirty'

interface Props {
  text: string
  map: ParseMap
  sid: number | null
  state: NodeState
  onSelect: (sid: number) => void
}

/** Max id chars shown in a crumb (minified pages carry KB-long ids). Exported for tests. */
export const MAX_CRUMB_ID_CHARS = 48

export function label(text: string | undefined, tag: string): string {
  if (!text) return tag
  const rawId = /\sid\s*=\s*["']([^"']+)["']/i.exec(text)?.[1]
  const id =
    rawId && rawId.length > MAX_CRUMB_ID_CHARS ? `${rawId.slice(0, MAX_CRUMB_ID_CHARS)}…` : rawId
  const cls = /\sclass\s*=\s*["']([^"']+)["']/i.exec(text)?.[1]
  return (
    tag + (id ? `#${id}` : '') + (cls ? `.${cls.trim().split(/\s+/).slice(0, 2).join('.')}` : '')
  )
}

/** Ancestor chain of the selected element; clicking a crumb selects that ancestor */
export function Breadcrumb({ text, map, sid, state, onSelect }: Props): ReactElement | null {
  const { t } = useI18n()
  if (sid === null) return <div className="crumbs crumbs-empty">{t('inspectHint')}</div>
  const current = map.bySid.get(sid)
  if (!current) return <div className="crumbs crumbs-empty">{t('inspectHint')}</div>
  const chain = [
    ...ancestorsOf(map, sid).filter((e) => !['html', 'head', 'body'].includes(e.tag)),
    current,
  ]
  return (
    <nav className="crumbs" aria-label={t('elementPath')}>
      {chain.map((e, i) => (
        <span key={`${e.sid}:${i}`} className="crumb-wrap">
          {i > 0 && (
            <span className="crumb-sep" aria-hidden>
              ›
            </span>
          )}
          <button
            type="button"
            className={`crumb${e.sid === sid ? ' current' : ''}`}
            aria-current={e.sid === sid ? true : undefined}
            onClick={() => onSelect(e.sid)}
            title={e.path}
          >
            {label(text.slice(e.startTag[0], e.startTag[1]), e.tag)}
          </button>
        </span>
      ))}
      {state !== 'static' && (
        <span
          className={`crumb-state ${state}`}
          title={t(state === 'dynamic' ? 'nodeDynamic' : 'nodeDirty')}
        >
          {t(state === 'dynamic' ? 'nodeDynamicShort' : 'nodeDirtyShort')}
        </span>
      )}
    </nav>
  )
}

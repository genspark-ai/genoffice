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

function label(text: string | undefined, tag: string): string {
  if (!text) return tag
  const id = /\sid\s*=\s*["']([^"']+)["']/i.exec(text)?.[1]
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
    <div className="crumbs" role="navigation" aria-label="element path">
      {chain.map((e, i) => (
        <span key={e.sid} className="crumb-wrap">
          {i > 0 && <span className="crumb-sep">›</span>}
          <button
            type="button"
            className={`crumb${e.sid === sid ? ' current' : ''}`}
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
    </div>
  )
}

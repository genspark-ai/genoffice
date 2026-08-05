import { memo } from 'react'
import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { collectHeadings } from '../editor/headings'
import { useI18n } from '../i18n/locale'

/** Word's navigation pane: heading outline with click-to-jump. Memoized on the doc — caret moves skip the outline walk. */
export const NavPane = memo(function NavPane({ editor, doc }: { editor: Editor; doc: PmNode }) {
  const { t } = useI18n()
  const headings = collectHeadings(doc)
  return (
    <aside className="nav-pane">
      <div className="nav-pane-title">{t('appNavTitle')}</div>
      <div className="nav-pane-list">
        {headings.map((h, i) => (
          <button
            key={`${h.pos}-${i}`}
            className={`nav-item nav-l${Math.min(h.level, 4)}`}
            title={h.text}
            onClick={() => {
              const dom = editor.view.nodeDOM(h.pos) as HTMLElement | null
              dom?.scrollIntoView({ behavior: 'smooth', block: 'start' })
            }}
          >
            {h.text}
          </button>
        ))}
        {headings.length === 0 && <div className="nav-empty">{t('appNavNoHeadings')}</div>}
      </div>
    </aside>
  )
})

import type { Editor } from '@tiptap/core'
import {
  isPageObjectSelected,
  pageObjects,
  selectPageObject,
  type PageObjectKind,
} from '../editor/page-objects'
import { toggleObjectSelection } from '../editor/object-multi-select'
import { useI18n, type StringKey } from '../i18n/locale'

const KIND_LABELS: Record<PageObjectKind, StringKey> = {
  picture: 'ribbonPicture',
  textBox: 'ribbonTextBox',
  shape: 'ribbonSelPaneShape',
}

/** Word's Selection Pane: the document's pictures, text boxes and shapes; a click selects one */
export function SelectionPane({ editor, onClose }: { editor: Editor; onClose: () => void }) {
  const { t } = useI18n()
  const objects = pageObjects(editor.state.doc)
  return (
    <aside className="comments-pane selection-pane">
      <div className="comments-pane-head">
        <span className="comments-pane-title">{t('ribbonSelectionPane')}</span>
        <button
          className="comments-pane-close"
          data-tip={t('appClose')}
          aria-label={t('appClose')}
          onClick={onClose}
        >
          ✕
        </button>
      </div>
      <div className="comments-pane-list" role="listbox" aria-label={t('ribbonSelectionPane')}>
        {objects.length === 0 && <div className="comments-empty">{t('ribbonSelPaneEmpty')}</div>}
        {objects.map((object) => {
          const selected = isPageObjectSelected(editor.state, object)
          return (
            <button
              key={`${object.pos}:${object.index ?? 'node'}`}
              type="button"
              role="option"
              aria-selected={selected}
              className={`selection-pane-row${selected ? ' active' : ''}`}
              // command-exempt: picks which object the selection is on, like a click on the page
              onClick={(event) => {
                // Shift or the command key adds a floating object to those selected, like on the page
                const adding = event.shiftKey || event.metaKey || event.ctrlKey
                if (!adding || !toggleObjectSelection(editor.view, object.pos))
                  selectPageObject(editor, object)
              }}
            >
              <span>
                {t(KIND_LABELS[object.kind])} {object.n}
              </span>
              {object.text && <span className="selection-pane-text">{object.text}</span>}
            </button>
          )
        })}
      </div>
    </aside>
  )
}

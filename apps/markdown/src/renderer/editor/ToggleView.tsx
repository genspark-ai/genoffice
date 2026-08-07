import { NodeViewContent, NodeViewWrapper } from '@tiptap/react'
import type { NodeViewProps } from '@tiptap/react'
import { t } from '../i18n/locale'

export function ToggleView({ node, updateAttributes, editor, getPos }: NodeViewProps) {
  const open: boolean = node.attrs.open !== false

  // fold state is display-only and never serializes — flip it outside the
  // undo history and tagged uiOnly so it neither dirties the document nor
  // becomes a Cmd+Z step
  const flipOpen = () => {
    const pos = getPos()
    if (typeof pos !== 'number') return
    const tr = editor.state.tr
      .setNodeMarkup(pos, undefined, { ...node.attrs, open: !open })
      .setMeta('addToHistory', false)
      .setMeta('uiOnly', true)
    editor.view.dispatch(tr)
  }

  return (
    <NodeViewWrapper className={`md-toggle ${open ? 'is-open' : 'is-closed'}`} data-toggle="">
      <div className="md-toggle-header" contentEditable={false}>
        <button
          type="button"
          className="md-toggle-caret"
          aria-expanded={open}
          onMouseDown={(e) => e.preventDefault()}
          onClick={flipOpen}
        >
          <svg width="12" height="12" viewBox="0 0 12 12" aria-hidden="true">
            <path d="M4 2l4 4-4 4" stroke="currentColor" strokeWidth="1.6" fill="none" />
          </svg>
        </button>
        <input
          className="md-toggle-summary"
          value={node.attrs.summary ?? ''}
          placeholder={t('toggleSummaryPlaceholder')}
          readOnly={!editor.isEditable}
          onChange={(e) => updateAttributes({ summary: e.target.value })}
        />
      </div>
      <NodeViewContent className="md-toggle-body" style={open ? undefined : { display: 'none' }} />
    </NodeViewWrapper>
  )
}

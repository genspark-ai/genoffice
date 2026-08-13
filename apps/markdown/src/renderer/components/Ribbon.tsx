import { memo } from 'react'
import type { Editor } from '@tiptap/react'

interface RibbonProps {
  editor: Editor
  activeTab: 'home' | 'insert' | 'view'
  onTabChange: (tab: 'home' | 'insert' | 'view') => void
  zoom: number
  onToggleDark?: () => void
}

function isActive(editor: Editor, name: string, attrs?: Record<string, unknown>): boolean {
  return editor.isActive(name, attrs)
}

export const Ribbon = memo(function Ribbon({
  editor,
  activeTab,
  onTabChange,
  zoom,
  onToggleDark,
}: RibbonProps) {
  return (
    <div className="ribbon">
      <div className="ribbon-tabs">
        {(['home', 'insert', 'view'] as const).map((tab) => (
          <div
            key={tab}
            className={`ribbon-tab ${activeTab === tab ? 'active' : ''}`}
            onClick={() => onTabChange(tab)}
          >
            {tab === 'home' ? 'Home' : tab === 'insert' ? 'Insert' : 'View'}
          </div>
        ))}
      </div>
      <div className="ribbon-content">
        {activeTab === 'home' && <HomeTab editor={editor} />}
        {activeTab === 'insert' && <InsertTab editor={editor} />}
        {activeTab === 'view' && <ViewTab zoom={zoom} onToggleDark={onToggleDark} />}
      </div>
    </div>
  )
})

function FormatBtn({ editor, mark, label }: { editor: Editor; mark: string; label: string }) {
  return (
    <button
      className={`ribbon-btn ${isActive(editor, mark) ? 'active' : ''}`}
      onClick={() => editor.chain().focus().toggleMark(mark).run()}
      title={label}
    >
      {label.charAt(0)}
    </button>
  )
}

function HeadingBtn({ editor, level, label }: { editor: Editor; level: 1 | 2 | 3; label: string }) {
  return (
    <button
      className={`ribbon-btn ${isActive(editor, 'heading', { level }) ? 'active' : ''}`}
      onClick={() => editor.chain().focus().toggleNode('heading', 'paragraph', { level }).run()}
      title={label}
    >
      H{level}
    </button>
  )
}

function HomeTab({ editor }: { editor: Editor }) {
  return (
    <>
      <div className="ribbon-group">
        <FormatBtn editor={editor} mark="bold" label="B" />
        <FormatBtn editor={editor} mark="italic" label="I" />
        <FormatBtn editor={editor} mark="strike" label="S" />
        <FormatBtn editor={editor} mark="underline" label="U" />
        <FormatBtn editor={editor} mark="code" label="<>" />
      </div>
      <div className="ribbon-group">
        <HeadingBtn editor={editor} level={1} label="H1" />
        <HeadingBtn editor={editor} level={2} label="H2" />
        <HeadingBtn editor={editor} level={3} label="H3" />
      </div>
      <div className="ribbon-group">
        <button
          className={`ribbon-btn ${isActive(editor, 'bulletList') ? 'active' : ''}`}
          onClick={() => editor.chain().focus().toggleList('bulletList', 'listItem').run()}
          title="Bullet List"
        >
          •
        </button>
        <button
          className={`ribbon-btn ${isActive(editor, 'orderedList') ? 'active' : ''}`}
          onClick={() => editor.chain().focus().toggleList('orderedList', 'listItem').run()}
          title="Ordered List"
        >
          1.
        </button>
        <button
          className={`ribbon-btn ${isActive(editor, 'taskList') ? 'active' : ''}`}
          onClick={() => editor.chain().focus().toggleList('taskList', 'taskItem').run()}
          title="Task List"
        >
          ☑
        </button>
      </div>
      <div className="ribbon-group">
        <button
          className={`ribbon-btn ${isActive(editor, 'blockquote') ? 'active' : ''}`}
          onClick={() => editor.chain().focus().toggleNode('blockquote', 'paragraph').run()}
          title="Blockquote"
        >
          "
        </button>
        <button
          className={`ribbon-btn ${isActive(editor, 'codeBlock') ? 'active' : ''}`}
          onClick={() => editor.chain().focus().toggleNode('codeBlock', 'paragraph').run()}
          title="Code Block"
        >
          {}
        </button>
        <button
          className="ribbon-btn"
          // @ts-expect-error TipTap duplicate module type mismatch in workspace monorepo
          onClick={() => editor.commands.setHorizontalRule()}
          title="Horizontal Rule"
        >
          —
        </button>
      </div>
      <div className="ribbon-group">
        <button
          className={`ribbon-btn ${isActive(editor, 'link') ? 'active' : ''}`}
          onClick={() => {
            const url = window.prompt('Enter URL:')
            // @ts-expect-error TipTap duplicate module type mismatch in workspace monorepo
            if (url) editor.commands.setLink({ href: url })
          }}
          title="Link"
        >
          🔗
        </button>
      </div>
      <div className="ribbon-group">
        <button
          className={`ribbon-btn ${editor.isActive({ textAlign: 'left' }) ? 'active' : ''}`}
          // @ts-expect-error TipTap duplicate module type mismatch in workspace monorepo
          onClick={() => editor.commands.setTextAlign('left')}
          title="Align Left"
        >
          ≡
        </button>
        <button
          className={`ribbon-btn ${editor.isActive({ textAlign: 'center' }) ? 'active' : ''}`}
          // @ts-expect-error TipTap duplicate module type mismatch in workspace monorepo
          onClick={() => editor.commands.setTextAlign('center')}
          title="Align Center"
        >
          ≡
        </button>
        <button
          className={`ribbon-btn ${editor.isActive({ textAlign: 'right' }) ? 'active' : ''}`}
          // @ts-expect-error TipTap duplicate module type mismatch in workspace monorepo
          onClick={() => editor.commands.setTextAlign('right')}
          title="Align Right"
        >
          ≡
        </button>
      </div>
    </>
  )
}

function InsertTab({ editor }: { editor: Editor }) {
  return (
    <>
      <div className="ribbon-group">
        <button
          className="ribbon-btn ribbon-btn-wide"
          onClick={() => {
            const url = window.prompt('Enter image URL:')
            // @ts-expect-error TipTap duplicate module type mismatch in workspace monorepo
            if (url) editor.commands.setImage({ src: url })
          }}
        >
          Image
        </button>
        <button
          className="ribbon-btn ribbon-btn-wide"
          onClick={() => {
            const url = window.prompt('Enter URL:')
            // @ts-expect-error TipTap duplicate module type mismatch in workspace monorepo
            if (url) editor.commands.setLink({ href: url })
          }}
        >
          Link
        </button>
      </div>
      <div className="ribbon-group">
        <button
          className="ribbon-btn ribbon-btn-wide"
          // @ts-expect-error TipTap duplicate module type mismatch in workspace monorepo
          onClick={() => editor.commands.setHorizontalRule()}
        >
          HR
        </button>
        <button
          className="ribbon-btn ribbon-btn-wide"
          // @ts-expect-error TipTap duplicate module type mismatch in workspace monorepo
          onClick={() => editor.commands.insertTable({ rows: 3, cols: 3, withHeaderRow: true })}
        >
          Table
        </button>
      </div>
      <div className="ribbon-group">
        <button
          className="ribbon-btn ribbon-btn-wide"
          onClick={() => editor.chain().focus().toggleNode('codeBlock', 'paragraph').run()}
        >
          Code Block
        </button>
        <button
          className="ribbon-btn ribbon-btn-wide"
          onClick={() => editor.chain().focus().toggleNode('blockquote', 'paragraph').run()}
        >
          Quote
        </button>
      </div>
    </>
  )
}

function ViewTab({ zoom, onToggleDark }: { zoom: number; onToggleDark?: () => void }) {
  return (
    <>
      <div className="ribbon-group">
        <span style={{ fontSize: 11, color: 'var(--text-secondary)', padding: '0 8px' }}>
          Zoom: {zoom}%
        </span>
      </div>
      {onToggleDark && (
        <div className="ribbon-group">
          <button className="ribbon-btn ribbon-btn-wide" onClick={onToggleDark}>
            Dark Mode
          </button>
        </div>
      )}
    </>
  )
}

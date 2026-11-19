import { useState } from 'react'
import type { ReactElement } from 'react'

export interface OutlineNode {
  title: string
  bold?: boolean
  italic?: boolean
  dest?: unknown
  url?: string
  items?: OutlineNode[]
}

/** Max outline nesting rendered: a hostile PDF can nest bookmarks thousands
 *  deep, and unbounded recursion would overflow the render stack. Deeper
 *  levels are dropped (their ancestors still render). */
export const MAX_OUTLINE_DEPTH = 32

function Item({
  node,
  depth,
  currentDest,
  onGo,
}: {
  node: OutlineNode
  depth: number
  currentDest?: unknown
  onGo: (n: OutlineNode) => void
}): ReactElement {
  const hasChildren = (node.items?.length ?? 0) > 0
  const isCurrent = currentDest != null && node.dest != null && node.dest === currentDest
  return (
    <>
      <button
        type="button"
        role="treeitem"
        aria-level={depth + 1}
        aria-expanded={hasChildren ? true : undefined}
        aria-current={isCurrent ? true : undefined}
        className="pdf-outline-item"
        style={{
          paddingLeft: 10 + depth * 14,
          fontWeight: node.bold ? 600 : 400,
          fontStyle: node.italic ? 'italic' : undefined,
        }}
        data-tip={node.title}
        onClick={() => onGo(node)}
      >
        {node.title}
      </button>
      {depth < MAX_OUTLINE_DEPTH &&
        node.items?.map((c, i) => (
          <Item key={i} node={c} depth={depth + 1} currentDest={currentDest} onGo={onGo} />
        ))}
    </>
  )
}

/** One node of the edited (pending) bookmark tree; pageIndex is 0-based */
export interface PendingBookmarkNode {
  title: string
  pageIndex: number
  top?: number | null
  children?: PendingBookmarkNode[]
}

/** Bookmark edit affordances over the pending tree: page-target rows with
    inline rename and delete; jumps go straight to the page. */
function PendingItem({
  node,
  path,
  depth,
  editable,
  labels,
  onGoToPage,
  onRename,
  onDelete,
}: {
  node: PendingBookmarkNode
  path: string
  depth: number
  /** rename / delete affordances shown; a committed (unsaved) tree lists read-only */
  editable: boolean
  labels: { rename: string; delete: string }
  onGoToPage: (pageIndex: number) => void
  onRename: (path: string, title: string) => void
  onDelete: (path: string) => void
}): ReactElement {
  const [editing, setEditing] = useState(false)
  const [draft, setDraft] = useState(node.title)
  const commit = (): void => {
    setEditing(false)
    const trimmed = draft.trim()
    if (trimmed !== '' && trimmed !== node.title) onRename(path, trimmed)
  }
  return (
    <>
      <div className="pdf-outline-item pdf-bookmark-row" style={{ paddingLeft: 10 + depth * 14 }}>
        {editing ? (
          <input
            autoFocus
            className="pdf-modal-input"
            value={draft}
            onChange={(e) => setDraft(e.target.value)}
            onBlur={commit}
            onKeyDown={(e) => {
              if (e.key === 'Enter') commit()
              if (e.key === 'Escape') setEditing(false)
            }}
          />
        ) : (
          <>
            <button
              type="button"
              className="pdf-bookmark-title"
              onClick={() => onGoToPage(node.pageIndex)}
              data-tip={`p. ${node.pageIndex + 1}`}
            >
              {node.title}
            </button>
            {editable && (
              <span className="pdf-bookmark-actions">
                <button
                  type="button"
                  className="rb-icon"
                  aria-label={labels.rename}
                  data-tip={labels.rename}
                  onClick={() => setEditing(true)}
                >
                  ✎
                </button>
                <button
                  type="button"
                  className="rb-icon"
                  aria-label={labels.delete}
                  data-tip={labels.delete}
                  onClick={() => onDelete(path)}
                >
                  ✕
                </button>
              </span>
            )}
          </>
        )}
      </div>
      {node.children?.map((child, i) => (
        <PendingItem
          key={`${path}.${i}`}
          node={child}
          path={`${path}.${i}`}
          depth={depth + 1}
          editable={editable}
          labels={labels}
          onGoToPage={onGoToPage}
          onRename={onRename}
          onDelete={onDelete}
        />
      ))}
    </>
  )
}

/** Outline (bookmark) tree: click jumps to internal destinations; url entries open external links.
    While a pending (edited) tree is handed in, it renders instead — with rename/delete affordances
    and an add button while editing, read-only once the edit is committed but not yet saved;
    jumps go straight to the page the entry targets. */
export function OutlinePanel({
  outline,
  note,
  label,
  emptyLabel,
  currentDest,
  onGoToDest,
  pending,
  editing = false,
  labels,
  onGoToPage,
  onAdd,
  onRename,
  onDelete,
}: {
  outline: OutlineNode[]
  /** Caption above the tree, e.g. when the tree was derived from headings */
  note?: string
  /** Accessible name for the tree; caller passes t('outline') (English fallback kept local) */
  label?: string
  /** Empty-state copy; caller passes t('searchNoResults') (English fallback kept local) */
  emptyLabel?: string
  /** Destination of the current location, when known; matching item gets aria-current */
  currentDest?: unknown
  onGoToDest: (dest: unknown) => void
  /** Edited tree taking over the pane: being edited, or committed and waiting for Save */
  pending?: PendingBookmarkNode[] | null
  /** Bookmark editing is on: the add button and the per-row affordances show */
  editing?: boolean
  /** Action labels for the edit affordances */
  labels?: { add: string; rename: string; delete: string }
  onGoToPage?: (pageIndex: number) => void
  onAdd?: () => void
  onRename?: (path: string, title: string) => void
  onDelete?: (path: string) => void
}): ReactElement {
  const onGo = (n: OutlineNode) => {
    if (n.url) window.open(n.url, '_blank')
    else if (n.dest != null) onGoToDest(n.dest)
  }
  if (
    pending !== null &&
    pending !== undefined &&
    labels &&
    onGoToPage &&
    onAdd &&
    onRename &&
    onDelete
  ) {
    return (
      <div className="pdf-outline" role="tree" aria-label={label ?? 'Outline'}>
        {editing && (
          <button type="button" className="pdf-bookmark-add" onClick={onAdd}>
            {labels.add}
          </button>
        )}
        {pending.length === 0 ? (
          <div className="pdf-outline-empty" role="status">
            {emptyLabel ?? 'No results'}
          </div>
        ) : (
          pending.map((node, i) => (
            <PendingItem
              key={String(i)}
              node={node}
              path={String(i)}
              depth={0}
              editable={editing}
              labels={labels}
              onGoToPage={onGoToPage}
              onRename={onRename}
              onDelete={onDelete}
            />
          ))
        )}
      </div>
    )
  }
  return (
    <div className="pdf-outline" role="tree" aria-label={label ?? 'Outline'}>
      {note && <div className="pdf-outline-note">{note}</div>}
      {outline.length === 0 ? (
        <div className="pdf-outline-empty" role="status">
          {emptyLabel ?? 'No results'}
        </div>
      ) : (
        outline.map((n, i) => (
          <Item key={i} node={n} depth={0} currentDest={currentDest} onGo={onGo} />
        ))
      )}
    </div>
  )
}

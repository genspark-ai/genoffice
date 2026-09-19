import type { ReactElement } from 'react'

export interface OutlineNode {
  title: string
  bold?: boolean
  italic?: boolean
  dest?: unknown
  url?: string
  items?: OutlineNode[]
}

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
      {node.items?.map((c, i) => (
        <Item key={i} node={c} depth={depth + 1} currentDest={currentDest} onGo={onGo} />
      ))}
    </>
  )
}

/** Outline (bookmark) tree: click jumps to internal destinations; url entries open external links */
export function OutlinePanel({
  outline,
  note,
  label,
  emptyLabel,
  currentDest,
  onGoToDest,
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
}): ReactElement {
  const onGo = (n: OutlineNode) => {
    if (n.url) window.open(n.url, '_blank')
    else if (n.dest != null) onGoToDest(n.dest)
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

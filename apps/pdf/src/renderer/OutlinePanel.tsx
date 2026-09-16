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
  onGo,
}: {
  node: OutlineNode
  depth: number
  onGo: (n: OutlineNode) => void
}): ReactElement {
  return (
    <>
      <button
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
        <Item key={i} node={c} depth={depth + 1} onGo={onGo} />
      ))}
    </>
  )
}

/** Outline (bookmark) tree: click jumps to internal destinations; url entries open external links */
export function OutlinePanel({
  outline,
  note,
  onGoToDest,
}: {
  outline: OutlineNode[]
  /** Caption above the tree, e.g. when the tree was derived from headings */
  note?: string
  onGoToDest: (dest: unknown) => void
}): ReactElement {
  const onGo = (n: OutlineNode) => {
    if (n.url) window.open(n.url, '_blank')
    else if (n.dest != null) onGoToDest(n.dest)
  }
  return (
    <div className="pdf-outline">
      {note && <div className="pdf-outline-note">{note}</div>}
      {outline.map((n, i) => (
        <Item key={i} node={n} depth={0} onGo={onGo} />
      ))}
    </div>
  )
}

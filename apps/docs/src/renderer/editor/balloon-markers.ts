/**
 * Show Markup → Balloons display widgets (All Markup only):
 * - "Show Revisions in Balloons": deleted text leaves the flow, so each
 *   deletion gets a thin marker at its spot (Word's connector foot) — the
 *   balloon leader anchors there and right-click opens the change's menu
 *   (the position is read from the DOM at click time; mapped widgets survive
 *   edits elsewhere without a rebuild).
 * - "Show All Revisions Inline": comments lose their balloons; a Word-style
 *   `[initials N]` tag after the commented range carries the thread as a
 *   tooltip and opens the Comments pane on click.
 */
import { Extension } from '@tiptap/core'
import { PluginKey } from '@tiptap/pm/state'
import { Decoration } from '@tiptap/pm/view'
import type { Node as PmNode } from '@tiptap/pm/model'
import type { CommentInfo } from '@genoffice/docx-engine'
import { blockDecorationPlugin, containsNode } from './local-edit'
import { revisionDisplayState } from './marks'
import { commentAnchors } from './comments'

export type InlineCommentTag = { initials: string; no: number; title: string }

/** open threads the inline tags describe; App refreshes it with the comment list */
export const inlineCommentState = { tags: new Map<string, InlineCommentTag>(), version: 0 }

export function commentInitials(c: Pick<CommentInfo, 'author' | 'initials'>): string {
  if (c.initials) return c.initials
  return c.author
    .split(/\s+/)
    .filter(Boolean)
    .map((w) => w[0]!.toUpperCase())
    .join('')
}

export function setInlineComments(comments: CommentInfo[]): void {
  const tags = new Map<string, InlineCommentTag>()
  let no = 0
  for (const c of comments) {
    if (c.parentId || c.done) continue
    no += 1
    const replies = comments.filter((r) => r.parentId === c.id)
    const title = [`${c.author}: ${c.text}`, ...replies.map((r) => `${r.author}: ${r.text}`)].join(
      '\n',
    )
    tags.set(c.id, { initials: commentInitials(c), no, title })
  }
  inlineCommentState.tags = tags
  inlineCommentState.version += 1
}

const hasMark = (name: string) => (n: PmNode) =>
  n.isText && n.marks.some((m) => m.type.name === name)

function deletionMarkers(doc: PmNode): Decoration[] {
  const decos: Decoration[] = []
  let prevEnd = -1
  doc.descendants((node, pos) => {
    if (!node.isText) return true
    const del = node.marks.find((m) => m.type.name === 'del')
    if (!del) return true
    if (pos !== prevEnd) {
      const author = typeof del.attrs.author === 'string' ? del.attrs.author : ''
      decos.push(
        Decoration.widget(
          pos,
          () => {
            const el = document.createElement('span')
            el.className = `doc-del-marker${del.attrs.move ? ' doc-move-from' : ''}`
            el.dataset.delAuthor = author
            return el
          },
          { side: -1, key: `del-marker-${author}-${del.attrs.move ? 'm' : ''}` },
        ),
      )
    }
    prevEnd = pos + node.nodeSize
    return true
  })
  return decos
}

function commentTags(doc: PmNode): Decoration[] {
  const decos: Decoration[] = []
  for (const a of commentAnchors(doc)) {
    const tag = inlineCommentState.tags.get(a.id)
    if (!tag) continue
    decos.push(
      Decoration.widget(
        a.to,
        () => {
          const el = document.createElement('span')
          el.className = 'doc-comment-marker'
          el.dataset.commentId = a.id
          el.title = tag.title
          el.textContent = `[${tag.initials}${tag.no}]`
          return el
        },
        { side: 1, key: `comment-tag-${a.id}-${tag.initials}${tag.no}` },
      ),
    )
  }
  return decos
}

const balloonMarkersKey = new PluginKey('balloonMarkers')

export const BalloonMarkersExtension = Extension.create({
  name: 'balloonMarkers',
  addProseMirrorPlugins() {
    return [
      blockDecorationPlugin({
        key: balloonMarkersKey,
        build(doc) {
          if (revisionDisplayState.mode !== 'all' || !revisionDisplayState.balloonsActive) return []
          if (revisionDisplayState.balloons === 'revisions') return deletionMarkers(doc)
          if (revisionDisplayState.balloons === 'inline') return commentTags(doc)
          return []
        },
        needsRecompute: (node) =>
          containsNode(node, hasMark('del')) || containsNode(node, hasMark('comment')),
        signature: () =>
          `${revisionDisplayState.mode}|${revisionDisplayState.balloons}|${revisionDisplayState.balloonsActive}|${inlineCommentState.version}`,
      }),
    ]
  },
})

/**
 * Anchor resolution for the add_comment tool: a block index plus an optional
 * exact text match inside that block, turned into a document range the
 * comment mark can cover. Matching runs over the same live text the block
 * list shows (pending tracked deletions skipped), so a span copied from
 * read_blocks resolves to the visible run.
 */
import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { CommentInfo } from '@genoffice/docx-engine'
import { liveText } from './doc-utils'

export interface CommentAnchorInput {
  blockIndex: unknown
  text?: unknown
  occurrence?: unknown
}

export type AnchorResolution = { from: number; to: number; excerpt: string } | { error: string }

interface Segment {
  textStart: number
  textEnd: number
  pos: number
  /** an atom leaf (formula, ruby): its stand-in text has no character positions */
  leaf: boolean
}

/** live text of the block with, per segment, the document position its characters start at */
function liveSegments(
  block: ProseMirrorNode,
  blockStart: number,
): { text: string; segments: Segment[] } {
  let text = ''
  const segments: Segment[] = []
  block.descendants((child, pos) => {
    if (child.marks.some((m) => m.type.name === 'del')) return false
    if (child.isText) {
      const s = child.text ?? ''
      segments.push({
        textStart: text.length,
        textEnd: text.length + s.length,
        pos: blockStart + pos,
        leaf: false,
      })
      text += s
    } else if (child.isLeaf) {
      const s = child.type.spec.leafText?.(child) ?? ''
      if (s)
        segments.push({
          textStart: text.length,
          textEnd: text.length + s.length,
          pos: blockStart + pos,
          leaf: true,
        })
      text += s
    }
    return true
  })
  return { text, segments }
}

/** first paragraph inside a container block that still has live text; its content start position */
function firstTextParagraph(
  block: ProseMirrorNode,
  blockStart: number,
): { node: ProseMirrorNode; start: number } | null {
  let found: { node: ProseMirrorNode; start: number } | null = null
  block.descendants((child, pos) => {
    if (found) return false
    if (child.isTextblock) {
      if (liveSegments(child, 0).segments.some((s) => !s.leaf)) {
        found = { node: child, start: blockStart + pos + 1 }
      }
      return false
    }
    return true
  })
  return found
}

/** document position of live-text offset `off` (end = true maps the exclusive end of a match) */
function posOfOffset(segments: Segment[], off: number, end = false): number | null {
  for (const seg of segments) {
    const inside = end
      ? off > seg.textStart && off <= seg.textEnd
      : off >= seg.textStart && off < seg.textEnd
    if (!inside) continue
    if (seg.leaf) return null
    return seg.pos + (off - seg.textStart)
  }
  return null
}

export function resolveCommentAnchor(editor: Editor, input: CommentAnchorInput): AnchorResolution {
  const doc = editor.state.doc
  const count = doc.childCount
  const index = Number(input.blockIndex)
  if (!Number.isInteger(index) || index < 0 || index >= count) {
    return {
      error: `blockIndex must be an integer between 0 and ${count - 1} (the document has ${count} blocks)`,
    }
  }
  const block = doc.child(index)
  const blockStart = doc.resolve(0).posAtIndex(index) + 1
  const noText = input.text === undefined || input.text === null || input.text === ''
  // a table or other container anchors to its first paragraph: one range per id,
  // Word keeps only one commentRangeStart/End pair per comment
  const anchorBlock = block.isTextblock
    ? { node: block, start: blockStart }
    : firstTextParagraph(block, blockStart)
  if (!anchorBlock) {
    return {
      error: `block ${index} (${block.type.name}) has no text to anchor a comment to; pick a block with text`,
    }
  }
  const { text, segments } = liveSegments(anchorBlock.node, anchorBlock.start)
  const runs = segments.filter((s) => !s.leaf)
  if (runs.length === 0) {
    return {
      error: `block ${index} (${block.type.name}) has no text to anchor a comment to; pick a block with text`,
    }
  }
  if (noText) {
    const last = runs[runs.length - 1]!
    return {
      from: runs[0]!.pos,
      to: last.pos + (last.textEnd - last.textStart),
      excerpt: liveText(anchorBlock.node),
    }
  }
  if (typeof input.text !== 'string') return { error: 'text must be a string' }
  if (!block.isTextblock) {
    return {
      error: `block ${index} is a ${block.type.name}; omit text to anchor the comment to its first paragraph`,
    }
  }
  const needle = input.text
  const hits: number[] = []
  for (let at = text.indexOf(needle); at !== -1; at = text.indexOf(needle, at + 1)) hits.push(at)
  if (hits.length === 0) {
    return {
      error: `block ${index} does not contain "${clipText(needle)}"; text must match exactly`,
    }
  }
  const occurrence = input.occurrence === undefined ? undefined : Number(input.occurrence)
  if (occurrence !== undefined && (!Number.isInteger(occurrence) || occurrence < 1)) {
    return { error: 'occurrence must be a positive integer (1 = first match)' }
  }
  if (hits.length > 1 && occurrence === undefined) {
    return {
      error: `"${clipText(needle)}" occurs ${hits.length} times in block ${index}; set occurrence (1-${hits.length}) to pick one`,
    }
  }
  const pick = occurrence ?? 1
  if (pick > hits.length) {
    return {
      error: `occurrence ${pick} is out of range: "${clipText(needle)}" occurs ${hits.length} time(s) in block ${index}`,
    }
  }
  const offset = hits[pick - 1]!
  const from = posOfOffset(segments, offset)
  const to = posOfOffset(segments, offset + needle.length, true)
  if (from === null || to === null) {
    return {
      error: `"${clipText(needle)}" overlaps an inline object (formula or ruby); anchor to plain text`,
    }
  }
  return { from, to, excerpt: needle }
}

/** Replies of a thread root, in file order. */
export function repliesOf(comments: CommentInfo[], id: string): CommentInfo[] {
  return comments.filter((c) => c.parentId === id)
}

function clipText(s: string, max = 60): string {
  return s.length > max ? `${s.slice(0, max - 1)}…` : s
}

export const clipExcerpt = (text: string): string => clipText(text.replace(/\s+/g, ' ').trim(), 80)

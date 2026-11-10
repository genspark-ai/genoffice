import type { Node } from '@tiptap/pm/model'
import { serializeDocText, type DocEnvelope } from './docText'

/** Experimental no-op round trips; edited documents still use the legacy serializer. */
export function roundTripMarkdownEnabled(): boolean {
  try {
    return localStorage.getItem('mdapp.experimentalRoundTrip') === '1'
  } catch {
    return false
  }
}

export interface MarkdownSourceSnapshot {
  readonly source: string
  readonly envelope: Readonly<DocEnvelope>
  readonly doc: Node
}

export function captureMarkdownSource(
  source: string,
  envelope: DocEnvelope,
  doc: Node,
): MarkdownSourceSnapshot {
  return { source, envelope: { ...envelope }, doc }
}

/** Compare only at serialization time, never on the typing path. */
export function serializeMarkdown(
  envelope: DocEnvelope,
  doc: Node,
  serializeBody: () => string,
  original?: MarkdownSourceSnapshot,
): string {
  if (
    original &&
    envelope.frontmatter === original.envelope.frontmatter &&
    envelope.eol === original.envelope.eol &&
    envelope.trailingNewline === original.envelope.trailingNewline &&
    envelope.bom === original.envelope.bom &&
    doc.eq(original.doc)
  ) {
    return original.source
  }
  return serializeDocText(envelope, serializeBody())
}

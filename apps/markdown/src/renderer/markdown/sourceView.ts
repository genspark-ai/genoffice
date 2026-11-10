import type { Editor } from '@tiptap/core'
import { parseDocText, stripLegacyFencedDivs, type DocEnvelope } from './docText'
import { buildSourceMap, type SourceMap } from './sourceSplice'

export interface AppliedSource {
  /** the envelope re-parsed from the edited text: frontmatter, EOL style, BOM, EOF newline */
  envelope: DocEnvelope
  /** pairing for the next save: the text just parsed against the document just built */
  sourceMap: SourceMap | null
}

/**
 * Push source-view edits back into the editor.
 *
 * The editor stays the single source of truth. A hand-edited text is re-parsed
 * into the document rather than written to the file, so Save As asset
 * relocation, the round-trip snapshot and every other editor consumer keep
 * working with no new plumbing — the save path is unchanged.
 *
 * The history step is suppressed on purpose: one setContent per keystroke
 * would otherwise give every character its own undo entry, and the pane's own
 * native undo already covers the pane.
 */
export function applySourceText(editor: Editor, text: string): AppliedSource {
  const envelope = parseDocText(text)
  const body = stripLegacyFencedDivs(envelope.body)
  editor.chain().setMeta('addToHistory', false).setContent(body, { contentType: 'markdown' }).run()
  return { envelope, sourceMap: buildSourceMap(editor, editor.state.doc, body) }
}

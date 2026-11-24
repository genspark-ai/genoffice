/**
 * Edit > Paste Special > Paste link. A copy made in a saved document leaves a
 * token on the clipboard beside its content; the document itself remembers
 * the copied range. Pasting the link asks that document, through the main
 * process, to bookmark the range, and puts the content in as the result of a
 * LINK field naming the file and the bookmark, which Edit > Links lists and
 * updates. Nothing about the file travels on the clipboard.
 */
import type { Editor } from '@tiptap/core'
import { LINK_SOURCE_PREFIX, wordLinkInstruction } from '@genoffice/docx-engine'
import { collectBookmarks } from './cross-reference'
import { insertAsLinkResult } from './link-fields'
import { pastePayload, type ClipboardPayload } from './paste-actions'

const TOKEN = /<meta\b[^>]*\bname="genoffice-link"[^>]*\bcontent="([\w-]+)"/i
const MARK = 'docRangeBookmark'

/** the token a copy made in a document of this app left in the clipboard markup */
export const linkTokenOf = (html: string): string | null => TOKEN.exec(html)?.[1] ?? null

interface Copied {
  token: string
  from: number
  to: number
  /** the bookmark the range was given, once a link was pasted from it */
  bookmark?: string
}

/** the next free name for the range a link is pasted from */
function freeLinkName(editor: Editor): string {
  const taken = new Set(collectBookmarks(editor.state.doc, true).map((b) => b.name))
  editor.state.doc.descendants((node) => {
    for (const mark of node.marks) if (mark.type.name === MARK) taken.add(String(mark.attrs.name))
    return true
  })
  let n = 1
  while (taken.has(`${LINK_SOURCE_PREFIX}${n}`)) n++
  return `${LINK_SOURCE_PREFIX}${n}`
}

/**
 * The document's side: every copy of a text selection is offered as a link
 * source while the document has a file, and a claim bookmarks its range.
 * Returns the teardown.
 */
export function offerCopiesAsLinkSource(editor: Editor, hasFile: () => boolean): () => void {
  let copied: Copied | null = null
  const withdraw = () => {
    if (copied) window.desktop.offerLinkSource(null)
    copied = null
  }
  const onCopy = (event: ClipboardEvent) => {
    const { from, to, empty } = editor.state.selection
    const html = event.clipboardData?.getData('text/html')
    if (event.type !== 'copy' || empty || !html || !hasFile()) return withdraw()
    const token = crypto.randomUUID()
    event.clipboardData!.setData(
      'text/html',
      `${html}<meta name="genoffice-link" content="${token}">`,
    )
    copied = { token, from, to }
    window.desktop.offerLinkSource(token)
  }
  const onTransaction = ({
    transaction,
  }: {
    transaction: { mapping: Editor['state']['tr']['mapping'] }
  }) => {
    if (!copied) return
    copied.from = transaction.mapping.map(copied.from, 1)
    copied.to = transaction.mapping.map(copied.to, -1)
  }
  const dom = editor.view.dom
  dom.addEventListener('copy', onCopy)
  dom.addEventListener('cut', onCopy)
  editor.on('transaction', onTransaction)
  const offMark = window.desktop.onLinkSourceMark((token) => {
    if (!copied || copied.token !== token || !hasFile()) {
      window.desktop.linkSourceMarked(token, null)
      return
    }
    const type = editor.schema.marks[MARK]
    if (!copied.bookmark && type && copied.to > copied.from) {
      const name = freeLinkName(editor)
      const tr = editor.state.tr.addMark(copied.from, copied.to, type.create({ name }))
      if (tr.docChanged) {
        editor.view.dispatch(tr)
        copied.bookmark = name
      }
    }
    window.desktop.linkSourceMarked(token, copied.bookmark ?? null)
  })
  return () => {
    dom.removeEventListener('copy', onCopy)
    dom.removeEventListener('cut', onCopy)
    editor.off('transaction', onTransaction)
    offMark()
    withdraw()
  }
}

/** whether what is on the clipboard can be pasted as a link: its source is a saved document still open */
export async function canPasteLink(payload: ClipboardPayload): Promise<boolean> {
  const token = linkTokenOf(payload.html)
  return token !== null && (await window.desktop.claimLinkSource(token, false)) !== null
}

/**
 * Pastes the clipboard as a link to where it was copied from: `plain` keeps
 * the text only. False when the source is gone or nothing went in.
 */
export async function pasteLink(
  editor: Editor,
  payload: ClipboardPayload,
  plain: boolean,
): Promise<boolean> {
  const token = linkTokenOf(payload.html)
  const claim = token ? await window.desktop.claimLinkSource(token, true) : null
  if (!claim?.bookmark) return false
  const instr = wordLinkInstruction(claim.path, claim.bookmark, plain ? 't' : 'r')
  const { leftOut, linked } = insertAsLinkResult(editor, instr, () =>
    pastePayload(editor, payload, plain ? 'text' : 'source') ? true : null,
  )
  return leftOut === true && linked
}

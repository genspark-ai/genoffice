/**
 * Citations inserted from Docs itself (Quick Cite, the library pane), without
 * Zotero's own citation dialog. The field is exactly what Zotero writes, an
 * `ADDIN ZOTERO_ITEM CSL_CITATION {…}` field carrying each item's URI and
 * CSL-JSON, so Zotero's Refresh, its bibliography and the Word plugin all
 * treat it as theirs. The text starts as the citation Zotero formatted for
 * each item; the Refresh that follows the insert lets Zotero format the
 * cluster and number the citations in the document's style.
 */
import type { Editor } from '@tiptap/core'
import type { ZoteroLocalCiteData, ZoteroLocalErrorCode, ZoteroLocalResult } from '../../shared/ipc'

export const CSL_CITATION_SCHEMA =
  'https://github.com/citation-style-language/schema/raw/master/csl-citation.json'

/** the style Zotero formats the document in (its ZOTERO_PREF data); apa until one is set */
export function documentStyleId(documentData: string): string {
  const id = /<style\b[^>]*\bid="([^"]+)"/.exec(documentData)?.[1] ?? ''
  const short = /\/styles\/([a-z0-9][a-z0-9-]*)$/.exec(id)?.[1]
  return short ?? 'apa'
}

/** the URI Zotero records for an item of a synced user library; none for an unsynced one */
export function zoteroItemUri(libraryId: number, key: string): string | null {
  return libraryId > 0 ? `http://zotero.org/users/${libraryId}/items/${key}` : null
}

/** a citation Zotero formatted as HTML (some styles italicize or wrap it) as plain text */
export function plainCitation(html: string): string {
  return (
    html
      .replace(/<[^>]*>/g, '')
      .replace(/&nbsp;/g, '\u00a0')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      // Zotero writes some characters by number: "Lee &#38; Park"
      .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
      .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(Number(dec)))
      .replace(/&amp;/g, '&')
      .trim()
  )
}

/** what Zotero itself shows in a citation it has not formatted yet */
export const CITATION_PLACEHOLDER = '{Citation}'
const NUMERIC_CITATION_RE = /^[[(]?\d+(?:[-–,]\s*\d+)*[\])]?$/

/**
 * One text for several citations, until Zotero formats the cluster:
 * "(Smith, 2020)" and "(Lee, 2021)" become "(Smith, 2020; Lee, 2021)".
 * Citations that do not share their brackets are joined as they are.
 */
export function clusterText(citations: string[]): string {
  const texts = citations.map((citation) => citation.trim()).filter(Boolean)
  // a numeric style numbers each item alone as [1]: only Zotero's Refresh knows the real
  // numbers, so the citation shows Zotero's own placeholder rather than wrong ones
  if (texts.length > 0 && texts.every((text) => NUMERIC_CITATION_RE.test(text)))
    return CITATION_PLACEHOLDER
  if (texts.length <= 1) return texts[0] ?? ''
  for (const [open, close] of [
    ['(', ')'],
    ['[', ']'],
  ] as const) {
    if (texts.every((text) => text.startsWith(open) && text.endsWith(close))) {
      return `${open}${texts.map((text) => text.slice(1, -1).trim()).join('; ')}${close}`
    }
  }
  return texts.join('; ')
}

function randomCitationId(): string {
  const alphabet = 'abcdefghijklmnopqrstuvwxyz0123456789'
  let id = ''
  for (let i = 0; i < 8; i++) id += alphabet[Math.floor(Math.random() * alphabet.length)]
  return id
}

/** the field instruction Zotero writes for a citation of these items */
export function citationInstruction(
  items: ZoteroLocalCiteData[],
  formatted: string,
  citationID: string = randomCitationId(),
): string {
  const payload = {
    citationID,
    properties: { formattedCitation: formatted, plainCitation: formatted, noteIndex: 0 },
    citationItems: items.map((item) => {
      const uri = zoteroItemUri(item.libraryId, item.key)
      return {
        id: item.itemData.id ?? item.key,
        ...(uri ? { uris: [uri] } : {}),
        itemData: item.itemData,
      }
    }),
    schema: CSL_CITATION_SCHEMA,
  }
  return `ADDIN ZOTERO_ITEM CSL_CITATION ${JSON.stringify(payload)}`
}

/** puts a citation field of these items at the selection; false when the document cannot hold one */
export function insertZoteroCitation(editor: Editor, items: ZoteroLocalCiteData[]): boolean {
  if (items.length === 0 || !editor.isEditable) return false
  const { state, view } = editor
  const type = state.schema.marks.instrField
  if (!type) return false
  const text =
    clusterText(items.map((item) => plainCitation(item.citation))) || CITATION_PLACEHOLDER
  // no field id yet: the integration assigns one when Zotero next reads the fields
  const mark = type.create({
    instr: citationInstruction(items, text),
    beginXml: null,
    fieldId: null,
    fieldPart: 'single',
  })
  view.dispatch(state.tr.replaceSelectionWith(state.schema.text(text, [mark]), false))
  view.focus()
  return true
}

/** the document holds at least one Zotero citation or bibliography field */
export function hasZoteroFields(editor: Editor): boolean {
  const type = editor.state.schema.marks.instrField
  if (!type) return false
  let found = false
  editor.state.doc.descendants((node) => {
    if (found) return false
    if (
      node.isText &&
      node.marks.some(
        (mark) => mark.type === type && /ZOTERO_(ITEM|BIBL)/.test(String(mark.attrs.instr ?? '')),
      )
    ) {
      found = true
    }
    return !found
  })
  return found
}

/** the message a failed local API call shows: how to turn it on, or that Zotero is not running */
export function zoteroLocalErrorKey(
  code: ZoteroLocalErrorCode,
): 'zoteroLocalApiOff' | 'zoteroConnectionError' | 'zoteroOperationError' {
  if (code === 'off') return 'zoteroLocalApiOff'
  if (code === 'unreachable') return 'zoteroConnectionError'
  return 'zoteroOperationError'
}

export interface CiteFromZoteroDeps {
  /** the document's Zotero preferences (ZOTERO_PREF), '' before any were set */
  documentData: string
  /** footnotes or endnotes hold Zotero fields, which a Refresh would drop from the bibliography */
  noteFields: boolean
  citeData(keys: string[], style: string): Promise<ZoteroLocalResult<ZoteroLocalCiteData[]>>
  refresh(): Promise<unknown>
}

export type CiteFromZoteroResult =
  | { ok: true; refreshed: boolean }
  | { ok: false; error: ReturnType<typeof zoteroLocalErrorKey> | 'zoteroCannotInsert' }

/**
 * Cites these items at the selection. A document Zotero already has preferences
 * for is then refreshed, so Zotero formats the cluster, numbers it and updates
 * the bibliography. A document without them keeps the citation as formatted in
 * the default style: a Refresh would open Zotero's Document Preferences first,
 * and the user can set them from the Zotero tab whenever they want.
 */
export async function citeFromZotero(
  editor: Editor,
  keys: string[],
  deps: CiteFromZoteroDeps,
): Promise<CiteFromZoteroResult> {
  if (keys.length === 0) return { ok: true, refreshed: false }
  if (!editor.isEditable) return { ok: false, error: 'zoteroCannotInsert' }
  const data = await deps.citeData(keys, documentStyleId(deps.documentData))
  if (!data.ok) return { ok: false, error: zoteroLocalErrorKey(data.errorCode) }
  if (!insertZoteroCitation(editor, data.value)) return { ok: false, error: 'zoteroCannotInsert' }
  if (!deps.documentData || deps.noteFields) return { ok: true, refreshed: false }
  try {
    await deps.refresh()
    return { ok: true, refreshed: true }
  } catch (error) {
    // the citation is in; Zotero formats it on the next Refresh
    console.error('Zotero refresh after Quick Cite failed', error)
    return { ok: true, refreshed: false }
  }
}

import type { Editor } from '@tiptap/core'
import type { Node as PmNode } from '@tiptap/pm/model'
import { Selection, TextSelection, type Transaction } from '@tiptap/pm/state'
import {
  blockLinkFields,
  breakBlockLinkField,
  docxBookmarkRange,
  docxRootNamespaces,
  editBlockLinkField,
  isWordDocumentLink,
  linkResultIsPlain,
  parseDocx,
  parseLinkInstruction,
  pictureLinkToken,
  retargetImageBlip,
  withLinkAutomatic,
  withLinkSource,
  withPictureLink,
  type Block,
  type LinkFieldInfo,
  type Run,
  type TextboxDisplay,
} from '@genoffice/docx-engine'
import type { LinkSourceResult } from '../../shared/ipc'
import { fetchDocBytes } from '../doc-bytes'
import type { HfPictureLinkRow } from '../hf-sections'
import { heldPictureToken } from './convert'
import type { StringKey } from '../i18n/locale'
import { insertFileBody, type InsertFileNotes, type InsertFileSource } from './insert-file'
import { isLinkBegin as isBegin, isLinkEnd as isEnd } from './link-mark-pairing'
import { imageSizeOf } from './picture-bytes'

/** one row of Edit > Links: content of the document that comes from another file */
export type DocLink =
  /** a link field whose result is ordinary content between its two markers */
  | {
      type: 'field'
      begin: number
      end: number
      instr: string
      locked: boolean
      info: LinkFieldInfo
    }
  /** a link field inside a block that is kept as XML */
  | {
      type: 'blockField'
      pos: number
      index: number
      instr: string
      locked: boolean
      info: LinkFieldInfo
    }
  /**
   * A picture that names the file it was inserted from: a picture block at
   * `pos`, or with `at` a picture in a line of text (the picture node itself
   * is at `pos`) or in a text box of the block at `pos`. With `hf` it is a picture
   * of a header or footer: the document has no place for it, and the strip takes the change.
   */
  | { type: 'picture'; pos: number; target: string; at?: PicturePlace; hf?: HfPictureLinkRow }

/** a linked picture of a header or footer as a row of the list */
export const hfPictureLink = (row: HfPictureLinkRow): DocLink => ({
  type: 'picture',
  pos: -1,
  target: row.target,
  hf: row,
})

/** where a linked picture that is not a block of its own sits */
export type PicturePlace =
  { in: 'line' } | { in: 'textbox'; box: number; para: number; run: number }

type RunImage = NonNullable<Run['image']>

/** the linked pictures inside the text boxes of a block */
function textboxPictures(node: PmNode): Array<{ at: PicturePlace; target: string }> {
  const boxes = node.attrs.textboxes as TextboxDisplay[] | null
  const out: Array<{ at: PicturePlace; target: string }> = []
  boxes?.forEach((box, b) =>
    box.paras.forEach((para, p) =>
      para.runs.forEach((run, r) => {
        const target = run.image?.linkTarget
        if (target) out.push({ at: { in: 'textbox', box: b, para: p, run: r }, target })
      }),
    ),
  )
  return out
}

/** the picture a link stands for, as its place holds it: the node's own attributes, or a run of a text box */
function runImageOf(doc: PmNode, link: Extract<DocLink, { type: 'picture' }>): RunImage | null {
  const node = doc.nodeAt(link.pos)
  if (!node || !link.at) return null
  if (link.at.in === 'line')
    return node.type.name === 'docInlineImage' && node.attrs.link === link.target
      ? {
          dataUrl: String(node.attrs.dataUrl ?? ''),
          xml: String(node.attrs.xml ?? ''),
          linkTarget: link.target,
          ...(node.attrs.widthPx ? { widthPx: Number(node.attrs.widthPx) } : {}),
          ...(node.attrs.heightPx ? { heightPx: Number(node.attrs.heightPx) } : {}),
        }
      : null
  const { box, para, run } = link.at
  const image = (node.attrs.textboxes as TextboxDisplay[] | null)?.[box]?.paras[para]?.runs[run]
    ?.image
  return image?.linkTarget === link.target ? image : null
}

/** the picture of a link with `next` in its place; false when the link no longer stands */
function putRunImage(
  tr: Transaction,
  link: Extract<DocLink, { type: 'picture' }>,
  next: RunImage,
): boolean {
  const node = tr.doc.nodeAt(link.pos)
  if (!node || !link.at || !runImageOf(tr.doc, link)) return false
  if (link.at.in === 'line') {
    tr.setNodeMarkup(link.pos, undefined, {
      ...node.attrs,
      dataUrl: next.dataUrl,
      xml: next.xml,
      link: next.linkTarget ?? null,
      widthPx: next.widthPx ?? node.attrs.widthPx,
      heightPx: next.heightPx ?? node.attrs.heightPx,
      ...(next.dataUrl !== node.attrs.dataUrl ? { vector: null } : {}),
    })
    return true
  }
  const { box, para, run } = link.at
  const boxes = (node.attrs.textboxes as TextboxDisplay[]).map((b, bi) =>
    bi !== box
      ? b
      : {
          ...b,
          paras: b.paras.map((p, pi) =>
            pi !== para
              ? p
              : { ...p, runs: p.runs.map((r, ri) => (ri !== run ? r : { ...r, image: next })) },
          ),
        },
  )
  tr.setNodeMarkup(link.pos, undefined, { ...node.attrs, textboxes: boxes })
  return true
}

/** a run picture naming `target` from now on, or no file with null; its drawing says the same */
function withRunImageLink(image: RunImage, target: string | null): RunImage {
  const { linkTarget: _was, ...rest } = image
  const xml = image.xml
    ? withPictureLink(image.xml, target === null ? null : pictureLinkToken(target))
    : image.xml
  return { ...rest, xml, ...(target === null ? {} : { linkTarget: target }) }
}

const originalsCache = new WeakMap<Block[], Map<number, string>>()
function originalsOf(blocks: Block[]): Map<number, string> {
  let map = originalsCache.get(blocks)
  if (!map) {
    map = new Map()
    for (const block of blocks)
      if (block.docxIndex != null && block.originalXml) map.set(block.docxIndex, block.originalXml)
    originalsCache.set(blocks, map)
  }
  return map
}

/** the XML a protected block saves as, as far as its link fields go */
function blockXml(node: PmNode, blocks: Block[]): string {
  if (typeof node.attrs.linkXml === 'string') return node.attrs.linkXml
  if (node.attrs.genXml) return String(node.attrs.genXml)
  return originalsOf(blocks).get(node.attrs.docxIndex as number) ?? ''
}

/** the document's links in reading order */
export function collectLinks(doc: PmNode, blocks: Block[]): DocLink[] {
  const links: DocLink[] = []
  const open: { pos: number; node: PmNode }[] = []
  doc.descendants((node, pos) => {
    if (node.type.name === 'docProtected') {
      if (typeof node.attrs.imageLink === 'string' && node.attrs.imageLink)
        links.push({ type: 'picture', pos, target: node.attrs.imageLink })
      for (const picture of textboxPictures(node)) links.push({ type: 'picture', pos, ...picture })
      for (const field of blockLinkFields(blockXml(node, blocks))) {
        const info = parseLinkInstruction(field.instr)
        if (info) links.push({ type: 'blockField', pos, ...field, info })
      }
      return false
    }
    if (node.type.name === 'docInlineImage') {
      if (typeof node.attrs.link === 'string' && node.attrs.link)
        links.push({ type: 'picture', pos, target: node.attrs.link, at: { in: 'line' } })
      return false
    }
    if (isBegin(node)) open.push({ pos, node })
    else if (isEnd(node)) {
      const from = open.pop()
      const instr = String(from?.node.attrs.linkBegin ?? '')
      const info = parseLinkInstruction(instr)
      if (from && info)
        links.push({
          type: 'field',
          begin: from.pos,
          end: pos,
          instr,
          locked: from.node.attrs.locked === true,
          info,
        })
    }
    return true
  })
  // stable: the pictures of one block keep the order they were read in
  return links.sort((a, b) => linkPos(a) - linkPos(b))
}

export const linkPos = (link: DocLink): number => (link.type === 'field' ? link.begin : link.pos)

const hasLinksCache = new WeakMap<Block[], WeakMap<PmNode, boolean>>()

/** whether Edit > Links has anything to list; a block's answer is kept across edits elsewhere */
export function hasLinks(doc: PmNode, blocks: Block[]): boolean {
  let cache = hasLinksCache.get(blocks)
  if (!cache) hasLinksCache.set(blocks, (cache = new WeakMap()))
  const known = cache.get(doc)
  if (known !== undefined) return known
  let found = false
  doc.forEach((node) => {
    if (found) return
    let own = cache.get(node)
    if (own === undefined) {
      own = false
      if (node.type.name === 'docProtected')
        own =
          Boolean(node.attrs.imageLink) ||
          textboxPictures(node).length > 0 ||
          blockLinkFields(blockXml(node, blocks)).length > 0
      else
        node.descendants((child) => {
          if (isBegin(child) || (child.type.name === 'docInlineImage' && child.attrs.link))
            own = true
          return !own
        })
      cache.set(node, own)
    }
    found = own
  })
  cache.set(doc, found)
  return found
}

/** the field takes its result from a Word document: INCLUDETEXT, or a LINK to a range of one */
export const readsDocument = (info: LinkFieldInfo): boolean =>
  info.kind === 'INCLUDETEXT' || isWordDocumentLink(info)

/** what the list shows as the kind of a link */
export type LinkTypeKey = 'document' | 'picture' | 'object'
export function linkTypeOf(link: DocLink): LinkTypeKey {
  if (link.type === 'picture') return 'picture'
  if (readsDocument(link.info)) return 'document'
  return link.info.kind === 'INCLUDEPICTURE' ? 'picture' : 'object'
}

export const linkSourceOf = (link: DocLink): string =>
  link.type === 'picture' ? link.target : link.info.source

/** Update Now can read this link's source again: a field that reads a document, or a linked picture */
export const linkUpdatable = (link: DocLink): boolean =>
  link.type === 'picture' || (link.type === 'field' && readsDocument(link.info))

function setBlockXml(tr: Transaction, pos: number, xml: string | null): boolean {
  const node = tr.doc.nodeAt(pos)
  if (!node || xml === null) return false
  tr.setNodeMarkup(pos, undefined, { ...node.attrs, linkXml: xml })
  return true
}

/** Break Link: the field becomes its result, the picture keeps only its own bytes */
export function breakLinks(editor: Editor, links: DocLink[], blocks: Block[]): number {
  const tr = editor.state.tr
  let done = 0
  // from the end: an earlier link's positions stay good
  for (const link of [...links].sort((a, b) => linkPos(b) - linkPos(a))) {
    if (link.type === 'field') {
      tr.delete(link.end, link.end + 1).delete(link.begin, link.begin + 1)
      done++
    } else if (link.type === 'picture') {
      const node = link.hf ? null : tr.doc.nodeAt(link.pos)
      if (!node) continue
      if (link.at) {
        const image = runImageOf(tr.doc, link)
        if (!image || !putRunImage(tr, link, withRunImageLink(image, null))) continue
      } else tr.setNodeMarkup(link.pos, undefined, { ...node.attrs, imageLink: null })
      done++
    } else {
      const node = tr.doc.nodeAt(link.pos)
      // a later field of the same block went first, so this one's place in it stands
      if (
        node &&
        setBlockXml(tr, link.pos, breakBlockLinkField(blockXml(node, blocks), link.index))
      )
        done++
    }
  }
  if (done > 0) editor.view.dispatch(tr)
  return done
}

/** a field's instruction and lock as given; false when the link is not a field */
function editFieldIn(
  tr: Transaction,
  link: DocLink,
  blocks: Block[],
  change: { instr?: string; locked?: boolean },
): boolean {
  if (link.type === 'field') {
    const node = tr.doc.nodeAt(link.begin)
    if (!node || !isBegin(node)) return false
    tr.setNodeMarkup(link.begin, undefined, {
      ...node.attrs,
      ...(change.instr !== undefined ? { linkBegin: change.instr } : {}),
      ...(change.locked !== undefined ? { locked: change.locked } : {}),
    })
  } else if (link.type === 'blockField') {
    const node = tr.doc.nodeAt(link.pos)
    if (!node) return false
    if (!setBlockXml(tr, link.pos, editBlockLinkField(blockXml(node, blocks), link.index, change)))
      return false
  } else return false
  return true
}

/** the change of each link in one step, undone as one; returns how many took it */
function editFields(
  editor: Editor,
  links: DocLink[],
  blocks: Block[],
  change: (link: DocLink) => { instr?: string; locked?: boolean } | null,
): number {
  const tr = editor.state.tr
  let done = 0
  for (const link of links) {
    const next = change(link)
    if (next && editFieldIn(tr, link, blocks, next)) done++
  }
  if (done > 0) editor.view.dispatch(tr)
  return done
}

const editField = (
  editor: Editor,
  link: DocLink,
  blocks: Block[],
  change: { instr?: string; locked?: boolean },
): boolean => editFields(editor, [link], blocks, () => change) > 0

export const setLinksLocked = (
  editor: Editor,
  links: DocLink[],
  blocks: Block[],
  locked: boolean,
): number => editFields(editor, links, blocks, () => ({ locked }))

export const setLinkLocked = (
  editor: Editor,
  link: DocLink,
  blocks: Block[],
  locked: boolean,
): boolean => setLinksLocked(editor, [link], blocks, locked) > 0

/** the `\a` switch of every LINK field among `links`, in one step */
export const setLinksAutomatic = (
  editor: Editor,
  links: DocLink[],
  blocks: Block[],
  automatic: boolean,
): number =>
  editFields(editor, links, blocks, (link) =>
    link.type === 'picture' || link.info.kind !== 'LINK'
      ? null
      : { instr: withLinkAutomatic(link.instr, automatic) },
  )

/** a LINK field's `\a` switch: whether Word refreshes it without being asked */
export function setLinkAutomatic(
  editor: Editor,
  link: DocLink,
  blocks: Block[],
  automatic: boolean,
): boolean {
  if (link.type === 'picture' || link.info.kind !== 'LINK') return false
  return editField(editor, link, blocks, { instr: withLinkAutomatic(link.instr, automatic) })
}

/** the link names another file from now on; what it shows changes at the next update */
export function setLinkSource(
  editor: Editor,
  link: DocLink,
  blocks: Block[],
  source: { path: string; url: string },
): boolean {
  if (link.type !== 'picture')
    return editField(editor, link, blocks, { instr: withLinkSource(link.instr, source.path) })
  if (link.hf) return false
  if (link.at) {
    const tr = editor.state.tr
    const image = runImageOf(tr.doc, link)
    if (!image || !putRunImage(tr, link, withRunImageLink(image, source.url))) return false
    editor.view.dispatch(tr)
    return true
  }
  const node = editor.state.doc.nodeAt(link.pos)
  if (!node) return false
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(link.pos, undefined, { ...node.attrs, imageLink: source.url }),
  )
  return true
}

export type LinkUpdateResult =
  | { status: 'updated'; leftOut: number }
  | { status: 'locked' | 'notUpdatable' | 'empty' }
  | { status: 'failed'; reason: Exclude<LinkSourceResult, { ok: true }>['reason'] | 'bookmark' }

/** a picture in a line of text or a text box takes the file's bytes where it stands: its drawing and width stay */
function updateRunPicture(
  editor: Editor,
  link: Extract<DocLink, { type: 'picture' }>,
  dataUrl: string,
  natural: { width: number; height: number },
): LinkUpdateResult {
  const tr = editor.state.tr
  const image = runImageOf(tr.doc, link)
  if (!image) return { status: 'failed', reason: 'failed' }
  if (image.dataUrl === dataUrl) return { status: 'updated', leftOut: 0 }
  const widthPx = image.widthPx || Math.min(natural.width, 620)
  const { vectorUrl: _old, ...kept } = image
  const next: RunImage = {
    ...kept,
    dataUrl,
    // in a line the save writes the new height into the drawing; a text box keeps the frame it has
    ...(link.at?.in === 'line'
      ? {
          widthPx,
          heightPx: Math.max(1, Math.round((widthPx * natural.height) / natural.width)),
        }
      : {}),
    // a picture not saved yet has no drawing: the save writes one from its bytes
    xml: image.xml && retargetImageBlip(image.xml, heldPictureToken()),
  }
  if (!putRunImage(tr, link, withRunImageLink(next, link.target)))
    return { status: 'failed', reason: 'failed' }
  editor.view.dispatch(tr)
  return { status: 'updated', leftOut: 0 }
}

type PictureRead = Extract<LinkSourceResult, { kind: 'picture' }>

/** the picture a link names, read again through the main process; or why it could not be read */
export async function readLinkedPicture(
  target: string,
): Promise<PictureRead | Extract<LinkUpdateResult, { status: 'failed' }>> {
  const read = await window.desktop.readLinkSource(target, 'picture')
  if (!read.ok) return { status: 'failed', reason: read.reason }
  return read.kind === 'picture' ? read : { status: 'failed', reason: 'unsupported' }
}

async function updatePicture(
  editor: Editor,
  link: Extract<DocLink, { type: 'picture' }>,
): Promise<LinkUpdateResult> {
  const { pos, target } = link
  if (link.hf) return { status: 'notUpdatable' }
  const read = await readLinkedPicture(target)
  if ('status' in read) return read
  const dataUrl = `data:${read.mime};base64,${read.base64}`
  const natural = await imageSizeOf(dataUrl).catch(() => null)
  if (link.at)
    return natural
      ? updateRunPicture(editor, link, dataUrl, natural)
      : { status: 'failed', reason: 'failed' }
  const node = editor.state.doc.nodeAt(pos)
  if (!natural || node?.type.name !== 'docProtected' || node.attrs.imageLink !== target)
    return { status: 'failed', reason: 'failed' }
  if (node.attrs.imageDataUrl === dataUrl) return { status: 'updated', leftOut: 0 }
  const widthPx = Number(node.attrs.imageWidthPx) || Math.min(natural.width, 620)
  const heightPx = Math.max(1, Math.round((widthPx * natural.height) / natural.width))
  const fromFile = node.attrs.docxIndex !== null && node.attrs.docxIndex !== undefined
  editor.view.dispatch(
    editor.state.tr.setNodeMarkup(pos, undefined, {
      ...node.attrs,
      imageDataUrl: dataUrl,
      imageVector: null,
      imageHeightPx: heightPx,
      imageCrop: null,
      imageFillRect: null,
      ...(fromFile
        ? { imageReplace: { base64: read.base64, mime: read.mime, linked: true } }
        : { genImage: { base64: read.base64, mime: read.mime, widthPx, heightPx } }),
    }),
  )
  return { status: 'updated', leftOut: 0 }
}

/** the field at `begin` as it stands now, after whatever happened while its source was read */
function fieldAt(
  editor: Editor,
  begin: number,
  instr: string,
): { begin: number; end: number } | null {
  const { doc } = editor.state
  const node = doc.nodeAt(begin)
  if (!node || !isBegin(node) || node.attrs.linkBegin !== instr) return null
  let depth = 0
  let end = -1
  doc.nodesBetween(begin, doc.content.size, (child, pos) => {
    if (end !== -1 || pos < begin) return end === -1
    if (isBegin(child)) depth++
    else if (isEnd(child) && --depth === 0) end = pos
    return end === -1
  })
  return end === -1 ? null : { begin, end }
}

/**
 * Update Now: the source is read again and takes the place of what the link
 * shows. A source that cannot be read leaves the document as it is.
 */
export async function updateLink(
  editor: Editor,
  link: DocLink,
  notes: InsertFileNotes,
): Promise<LinkUpdateResult> {
  if (link.type === 'picture') return updatePicture(editor, link)
  if (link.type !== 'field' || !readsDocument(link.info)) return { status: 'notUpdatable' }
  if (link.locked) return { status: 'locked' }
  const read = await window.desktop.readLinkSource(link.info.source, 'document')
  if (!read.ok) return { status: 'failed', reason: read.reason }
  if (read.kind !== 'document') return { status: 'failed', reason: 'unsupported' }
  try {
    let bytes = await fetchDocBytes(read.dataUrl)
    let inline = false
    if (link.info.item) {
      const range = await docxBookmarkRange(bytes, link.info.item)
      if (!range) return { status: 'failed', reason: 'bookmark' }
      bytes = range.bytes
      inline = range.inline
    }
    const namespaces = await docxRootNamespaces(bytes)
    const parsed = await parseDocx(bytes)
    const now = fieldAt(editor, link.begin, link.instr)
    if (!now) return { status: 'failed', reason: 'failed' }
    const { doc } = editor.state
    editor.view.dispatch(
      editor.state.tr.setSelection(
        TextSelection.between(doc.resolve(now.begin + 1), doc.resolve(now.end)),
      ),
    )
    const plain = link.info.kind === 'LINK' && linkResultIsPlain(link.instr)
    const leftOut = insertFileBody(editor, { parsed, namespaces, inline, plain }, notes)
    return leftOut === null ? { status: 'empty' } : { status: 'updated', leftOut }
  } catch {
    return { status: 'failed', reason: 'failed' }
  }
}

/**
 * Insert > File with Link to File: the body goes in as it does without the
 * link, and what went in becomes the result of a field with this instruction,
 * in the same undo step. Returns what the insertion returns, and whether the
 * field could be put around it.
 */
export function insertLinkedFileBody(
  editor: Editor,
  source: InsertFileSource,
  notes: InsertFileNotes | undefined,
  instr: string,
): { leftOut: number | null; linked: boolean } {
  return insertAsLinkResult(editor, instr, () => insertFileBody(editor, source, notes))
}

/**
 * What `insert` puts into the document becomes the result of a field with
 * this instruction. `insert` returns null when it put nothing in.
 */
export function insertAsLinkResult<T>(
  editor: Editor,
  instr: string,
  insert: () => T | null,
): { leftOut: T | null; linked: boolean } {
  // where the paste put content: the range its transaction changed, followed through the
  // transactions that come after it (renumbering, list fixes), which touch other places too
  let from = Infinity
  let to = -Infinity
  const follow = ({ transaction }: { transaction: Transaction }) => {
    if (!transaction.docChanged) return
    const first = to < from
    transaction.mapping.maps.forEach((map) => {
      if (to >= from) {
        from = map.map(from, -1)
        to = map.map(to, 1)
      }
      if (first)
        map.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
          from = Math.min(from, newStart)
          to = Math.max(to, newEnd)
        })
    })
  }
  editor.on('transaction', follow)
  let leftOut: T | null
  try {
    leftOut = insert()
  } finally {
    editor.off('transaction', follow)
  }
  if (leftOut === null || to < from) return { leftOut, linked: false }
  const { state } = editor
  const type = state.schema.nodes.docXeMark
  // a marker stands in text: where the content begins or ends between blocks, the opening
  // one goes to the end of the paragraph before and the closing one to the start of the next
  const inText = (pos: number, prefer: 1 | -1): number | null => {
    const $pos = state.doc.resolve(pos)
    if ($pos.parent.isTextblock) return pos
    for (const dir of [prefer, -prefer] as const) {
      const near = Selection.near($pos, dir)
      const { $from } = near
      const sameContainer =
        $from.depth === $pos.depth + 1 && $from.start($pos.depth) === $pos.start($pos.depth)
      if (near instanceof TextSelection && sameContainer) return near.from
    }
    return null
  }
  const start = inText(from, -1)
  const stop = inText(to, 1)
  if (!type || start === null || stop === null || stop < start) return { leftOut, linked: false }
  const tr = state.tr
    .insert(stop, type.create({ linkEnd: true }))
    .insert(start, type.create({ linkBegin: ` ${instr} ` }))
  editor.view.dispatch(tr.setMeta('appendedTransaction', tr))
  return { leftOut, linked: true }
}

/** what a link shows after an update was asked for */
export function linkStatusKey(result: LinkUpdateResult): StringKey {
  if (result.status === 'updated') return 'ribbonLinksUpdated'
  if (result.status === 'locked') return 'ribbonLinksIsLocked'
  if (result.status === 'notUpdatable') return 'ribbonLinksNotHere'
  if (result.status !== 'failed') return 'ribbonInsertFileEmpty'
  if (result.reason === 'missing') return 'ribbonLinksMissing'
  return result.reason === 'bookmark' ? 'ribbonLinksNoBookmark' : 'ribbonLinksUnreadable'
}

/**
 * F9 on linked text: the INCLUDETEXT fields the selection touches are read
 * again, the outermost of nested ones, the last first so the others keep
 * their places.
 */
export async function updateLinksAtSelection(
  editor: Editor,
  blocks: Block[],
  notes: InsertFileNotes,
): Promise<LinkUpdateResult[]> {
  const { from, to } = editor.state.selection
  type Field = Extract<DocLink, { type: 'field' }>
  const touched = collectLinks(editor.state.doc, blocks).filter(
    (link): link is Field =>
      link.type === 'field' && readsDocument(link.info) && link.begin <= to && link.end >= from,
  )
  const outer = touched.filter(
    (link) => !touched.some((other) => other.begin < link.begin && other.end > link.end),
  )
  const results: LinkUpdateResult[] = []
  for (const link of outer.reverse()) results.push(await updateLink(editor, link, notes))
  return results
}

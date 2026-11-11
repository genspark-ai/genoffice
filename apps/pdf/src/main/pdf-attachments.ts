import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFRawStream,
  PDFString,
} from 'pdf-lib'
import type { PDFObject } from 'pdf-lib'
import type { PdfAttachmentOps } from '../shared/ipc'

/**
 * Embedded files (§7.11): the catalog's /Names → /EmbeddedFiles name tree maps
 * names to Filespec dictionaries, each pointing at the file's bytes through
 * /EF. Only a flat tree (/Names array, no /Kids) is edited — a /Kids hierarchy
 * fails the whole attachments edit rather than being rewritten by
 * approximation, because a subtly wrong tree silently loses files in other
 * readers. Removals leave the Filespec objects in the file; the tree is what
 * readers enumerate, and unreferenced objects are invisible to them.
 */

export class AttachmentsError extends Error {}

const EMBEDDED_FILES = PDFName.of('EmbeddedFiles')

/** the catalog's /EmbeddedFiles node, or null when the document has none */
function embeddedTree(doc: PDFDocument): PDFDict | null {
  const names = doc.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)
  if (!names) return null
  return names.lookupMaybe(EMBEDDED_FILES, PDFDict) ?? null
}

function decodeText(value: unknown): string {
  // Most producers write /F and /UF as literal strings; PDFString.toString()
  // returns the value with its parentheses, which is not a name.
  return value instanceof PDFHexString || value instanceof PDFString
    ? value.decodeText()
    : String(value)
}

/** label → the tree slot of one attachment (its name object and filespec value) */
interface Slot {
  nameObject: unknown
  filespec: unknown
}

function existingSlots(doc: PDFDocument, tree: PDFDict): Map<string, Slot> {
  const out = new Map<string, Slot>()
  const array = tree.lookupMaybe(PDFName.of('Names'), PDFArray)
  if (!array) return out
  for (let i = 0; i + 1 < array.size(); i += 2) {
    const spec = doc.context.lookup(array.get(i + 1))
    if (!(spec instanceof PDFDict)) continue
    // The name-tree key is what pdf.js reports as the attachment's id — keying
    // by /UF or /F drifts the moment a producer writes them differently.
    const key = decodeText(array.get(i))
    if (key) out.set(key, { nameObject: array.get(i), filespec: array.get(i + 1) })
  }
  return out
}

/**
 * Apply add/remove ops to the embedded-files tree. Removals run first; adds
 * then claim their names, and any surviving collision fails the edit. The
 * rebuilt /Names stays sorted by name, like producers write it.
 */
export function applyAttachments(doc: PDFDocument, ops: PdfAttachmentOps): void {
  const tree = embeddedTree(doc)
  if (tree && tree.get(PDFName.of('Kids')) !== undefined) {
    throw new AttachmentsError(
      'the embedded-files name tree uses /Kids; refusing to edit it by approximation',
    )
  }
  const slots = tree ? existingSlots(doc, tree) : new Map<string, Slot>()
  for (const name of ops.remove ?? []) {
    if (!slots.delete(name)) throw new AttachmentsError(`attachment not found: ${name}`)
  }
  for (const add of ops.add ?? []) {
    if (slots.has(add.name)) throw new AttachmentsError(`attachment name already used: ${add.name}`)
    slots.set(add.name, {
      nameObject: PDFHexString.fromText(add.name),
      filespec: filespecRef(doc, add),
    })
  }
  writeTree(doc, tree, slots)
}

function filespecRef(
  doc: PDFDocument,
  add: { name: string; bytesBase64: string; description?: string },
): unknown {
  const context = doc.context
  const bytes = Buffer.from(add.bytesBase64, 'base64')
  const streamDict = context.obj({
    Type: 'EmbeddedFile',
    Subtype: 'application/octet-stream',
    Params: context.obj({ Size: bytes.length }),
    Length: bytes.length,
  })
  const streamRef = context.nextRef()
  context.assign(streamRef, PDFRawStream.of(streamDict, new Uint8Array(bytes)))
  // context.obj maps a bare string to a NAME; the file's name entries are strings
  const name = PDFHexString.fromText(add.name)
  const desc = add.description ? PDFHexString.fromText(add.description) : undefined
  const spec = context.obj({
    Type: 'Filespec',
    F: name,
    UF: name,
    ...(desc ? { Desc: desc } : {}),
    EF: context.obj({ F: streamRef, UF: streamRef }),
  })
  return context.register(spec)
}

function writeTree(doc: PDFDocument, tree: PDFDict | null, slots: Map<string, Slot>): void {
  const context = doc.context
  const flat: PDFObject[] = []
  for (const [, slot] of [...slots.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1))) {
    flat.push(slot.nameObject as PDFObject, slot.filespec as PDFObject)
  }
  const array = context.obj(flat)
  if (tree) {
    tree.set(PDFName.of('Names'), array)
    tree.delete(PDFName.of('Limits'))
    return
  }
  const created = context.register(context.obj({ Names: array }))
  const catalogNames = doc.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)
  if (catalogNames) {
    catalogNames.set(EMBEDDED_FILES, created)
    return
  }
  doc.catalog.set(PDFName.of('Names'), context.register(context.obj({ EmbeddedFiles: created })))
}

/** labels a reader enumerates for this document, in tree order */
export function attachmentLabels(doc: PDFDocument): string[] {
  const tree = embeddedTree(doc)
  if (!tree) return []
  const out: string[] = []
  const array = tree.lookupMaybe(PDFName.of('Names'), PDFArray)
  if (!array) return out
  for (let i = 0; i + 1 < array.size(); i += 2) {
    const spec = doc.context.lookup(array.get(i + 1))
    if (!(spec instanceof PDFDict)) continue
    const key = decodeText(array.get(i))
    if (key) out.push(key)
  }
  return out
}

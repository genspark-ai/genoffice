import {
  type ColorScheme,
  type CustomPath,
  DEFAULT_SCHEME,
  type EscherShape,
  ESCHER,
  FSP,
  OPT,
  PRESET_BY_SHAPE_TYPE,
  type PictureMedia,
  PictureStore,
  type Rect,
  arcGeometry,
  boolProp,
  fixedProp,
  freeformPath,
  officeArtColor,
  parseDrawing,
} from './escher'
import { patternMaskOf, patternPaletteOf, patternPreset } from './patterns'
import { LegacyPptError } from './errors'
import {
  PPT_TYPES,
  type PptRecord,
  decodeTextBytes,
  decodeTextChars,
  isRecordContainer,
  payloadOf,
  rawTextBytes,
  rawTextChars,
  walkRecords,
  walkRecordsLenient,
} from './records'
import { metroParagraphIndents } from './metro'
import {
  type AutoNumber,
  type LevelStyle,
  type ResolvedParagraph,
  type ResolvedRun,
  type TextRuler,
  type StyleRuns,
  levelsForText,
  parseFontCollection,
  parseMasterStyle,
  parseStyleRuns,
  parseStyleTextProp9,
  parseTextRuler,
  resolveText,
} from './text-style'

/** Caps: fail fast, never walk an absurd deck. Input/stream caps live in the shared OLE2 opener. */
const MAX_SLIDES = 4096
/** Deepest record nesting walked: a real slide's text sits about six levels down. */
const MAX_RECORD_DEPTH = 64
/** One change-directory chain (edits or persist blocks) longer than this is damaged, not a deck. */
const MAX_EDIT_CHAIN = 1024
/** A single slide with more text sequences than this is not a real deck. */
const MAX_TEXT_SEQUENCES = 100_000
const MAX_SHAPES_PER_SLIDE = 20_000
const MAX_GROUP_DEPTH = 32

/**
 * The PowerPoint slide surface is measured in master units: 576 per inch, so
 * a 10 × 7.5 in deck is 5760 × 4320 (not the 576000 × 432000 a
 * hundredth-of-an-inch surface would use).
 */
const DEFAULT_SLIDE_WIDTH = 5760
const DEFAULT_SLIDE_HEIGHT = 4320
const MAX_SLIDE_MASTER_UNITS = 1_000_000
export const EMU_PER_MASTER_UNIT = 914400 / 576

const CURRENT_USER_TOKEN = 0xe391c05f
/** CurrentUserAtom headerToken of a password-protected (encrypted) deck. */
export const ENCRYPTED_CURRENT_USER_TOKEN = 0xf3d1c4df

export interface PptSlideText {
  /** the TextHeaderEnum of the sequence (0 title, 1 body, 2 notes, 4 other…) */
  kind: number
  text: string
}

export type PptFill =
  | { kind: 'solid'; color: string; alpha?: number }
  | {
      kind: 'gradient'
      stops: Array<{ pos: number; color: string }>
      angle: number
      /** a shade from the focus rectangle out (shadeCenter / shadeShape) instead of linear */
      radial?: { path: 'rect' | 'shape'; l: number; t: number; r: number; b: number }
    }
  | { kind: 'picture'; pib: number; tile?: boolean }
  | { kind: 'pattern'; preset: string; fg: string; bg: string }
  | { kind: 'background' }

export interface PptLine {
  color: string
  /** EMU */
  width: number
  dash?: string
  compound?: string
}

export interface PptTextBody {
  kind: number
  paragraphs: ResolvedParagraph[]
  /** EMU: left, top, right, bottom */
  insets: [number, number, number, number]
  anchor: 't' | 'ctr' | 'b'
  anchorCenter: boolean
  wrap: boolean
  autofit: boolean
  vertical?: 'vert' | 'eaVert'
}

interface PptShapeBase {
  /** EMU */
  x: number
  y: number
  cx: number
  cy: number
  /** degrees clockwise */
  rotation: number
  flipH: boolean
  flipV: boolean
}

export interface PptShape extends PptShapeBase {
  type: 'shape'
  preset?: string
  path?: CustomPath
  /** preset adjust values (adj1 → 60000ths of a degree for an arc) */
  adjust?: Record<string, number>
  fill?: PptFill
  line?: PptLine
  text?: PptTextBody
}

export interface PptPicture extends PptShapeBase {
  type: 'picture'
  pib: number
  line?: PptLine
  /** thousandths of a percent cropped from each side */
  crop?: { l: number; t: number; r: number; b: number }
  /** RRGGBB the picture shows through (pictureTransparent) */
  transparent?: string
}

export type PptAnyShape = PptShape | PptPicture

export interface PptSlide {
  textSequences: PptSlideText[]
  shapes: PptAnyShape[]
  background?: PptFill
}

export interface PptPresentation {
  slides: PptSlide[]
  /** slide surface in master units (576 per inch) */
  slideWidth: number
  slideHeight: number
  pictures: PictureStore
}

/**
 * The change directory: persistId → stream offset. PowerPoint's stream holds
 * one generation per save; the live UserEditAtom (named by the CurrentUser
 * stream) starts the walk, and each generation's PersistDirectoryAtom maps
 * the persist ids it wrote to their stream offsets. A sequential walk of the
 * stream would read stale generations.
 */
export interface PersistChain {
  directory: Map<number, number>
  /** the DocumentContainer's persist id, when the chain names one */
  docPersistId: number | null
}

/** The CurrentUserAtom's headerToken and the offset of the live UserEditAtom. */
export interface CurrentUserInfo {
  headerToken: number
  editOffset: number | null
}

/**
 * The CurrentUserAtom of the "Current User" stream. headerToken distinguishes
 * a plain deck (0xE391C05F) from a password-protected one (0xF3D1C4DF, checked
 * by the container opener); offsetToCurrentEdit follows the size and
 * headerToken fields.
 */
export function currentUserInfo(currentUser: Uint8Array): CurrentUserInfo | null {
  for (const record of walkRecords(currentUser)) {
    if (record.type !== PPT_TYPES.currentUserAtom) continue
    const payload = payloadOf(currentUser, record)
    if (payload.length < 12) return { headerToken: 0, editOffset: null }
    const view = viewOf(payload)
    return {
      headerToken: view.getUint32(4, true),
      editOffset: view.getUint32(8, true) || null,
    }
  }
  return null
}

/**
 * The live change directory. The UserEditAtom the CurrentUserAtom points at
 * names (a) its persist directory (a PersistDirectoryAtom: 4-byte
 * persistId:20|cPersist:12 entries, then cPersist 4-byte offsets) and (b) the
 * previous generation through offsetLastEdit. Newer entries win: a persist id
 * the current generation wrote keeps its current offset even when an older
 * directory also lists it.
 */
export function buildPersistChain(
  document: Uint8Array,
  currentUser: Uint8Array | null,
): PersistChain {
  const directory = new Map<number, number>()
  let docPersistId: number | null = null
  const info = currentUser ? currentUserInfo(currentUser) : null
  if (!info || info.headerToken !== CURRENT_USER_TOKEN || info.editOffset === null) {
    return { directory, docPersistId }
  }
  let editOffset: number | null = info.editOffset
  for (let edits = 0; editOffset !== null && edits < MAX_EDIT_CHAIN; edits++) {
    const edit = recordAt(document, editOffset)
    if (!edit || edit.type !== PPT_TYPES.userEditAtom || edit.length < 28) {
      // the live edit is the one piece of the chain that must be readable
      if (edits === 0) throw new LegacyPptError('bad_persist', `no UserEditAtom at ${editOffset}`)
      break
    }
    const view = viewOf(payloadOf(document, edit))
    if (edits === 0) {
      const docRef = view.getUint32(16, true)
      if (docRef > 0) docPersistId = docRef
    }
    readPersistDirectory(document, view.getUint32(12, true), directory)
    editOffset = view.getUint32(8, true) || null
  }
  return { directory, docPersistId }
}

/** One generation's PersistDirectoryAtom chain; entries already in the map stay. */
function readPersistDirectory(
  document: Uint8Array,
  offset: number,
  directory: Map<number, number>,
): void {
  let at = offset
  for (let blocks = 0; blocks < MAX_EDIT_CHAIN; blocks++) {
    const record = recordAt(document, at)
    if (!record || record.type !== PPT_TYPES.persistPtrIncrementalBlock) return
    const payload = payloadOf(document, record)
    const view = viewOf(payload)
    let cursor = 0
    while (cursor + 4 <= payload.length) {
      const header = view.getUint32(cursor, true)
      cursor += 4
      const persistId = header & 0xfffff
      const count = header >>> 20
      if (count === 0 || cursor + count * 4 > payload.length) {
        throw new LegacyPptError('bad_persist', 'a persist directory entry runs off its block')
      }
      for (let index = 0; index < count; index++) {
        const streamOffset = view.getUint32(cursor, true)
        cursor += 4
        const id = persistId + index
        // offset 0 is a legal stream position (the DocumentContainer can open
        // the stream); only a past-the-end offset is garbage
        if (streamOffset < document.length && !directory.has(id)) {
          directory.set(id, streamOffset)
        }
      }
    }
    // the continuation blocks of one generation sit back to back
    at = record.offset + record.length
  }
}

/* ---------- document-wide context ---------- */

interface OutlineEntry {
  kind: number
  raw: string
  styles: StyleRuns | null
  ruler: TextRuler | null
  fields: TextField[]
  autoNumbers?: Array<AutoNumber | null>
}

/** a MetaCharacter atom: the field a '*' placeholder character stands for */
interface TextField {
  position: number
  type: 'slidenum' | 'datetime1' | 'header' | 'footer'
}

const META_CHAR_TYPES: Record<number, TextField['type']> = {
  0x0fd8: 'slidenum',
  0x0ff7: 'datetime1',
  0x0ff8: 'datetime1',
  0x1015: 'datetime1',
  0x0ff9: 'header',
  0x0ffa: 'footer',
}

function metaCharField(stream: Uint8Array, record: PptRecord): TextField | null {
  const type = META_CHAR_TYPES[record.type]
  if (!type || record.length < 4) return null
  return { position: viewOf(payloadOf(stream, record)).getInt32(0, true), type }
}

interface SlideAtomInfo {
  masterIdRef: number
  masterObjects: boolean
  masterScheme: boolean
  masterBackground: boolean
}

interface MasterInfo {
  shapes: EscherShape[]
  scheme: ColorScheme | null
  /** TxMasterStyleAtom instance → levels */
  styles: Map<number, LevelStyle[]>
  /** placeholder type → the master's placeholder shape: slide placeholders inherit its anchor and properties */
  placeholders: Map<number, EscherShape>
  background?: PptFill
  atom: SlideAtomInfo | null
}

interface DocContext {
  stream: Uint8Array
  fonts: string[]
  docStyles: Map<number, LevelStyle[]>
  pictures: PictureStore
  masters: Map<number, MasterInfo>
  /** masterId (what a SlideAtom.masterIdRef names) → persist id of the master container */
  masterPersistIds: Map<number, number>
  /** master ids whose parse is in progress: a masterIdRef cycle stops here */
  mastersInProgress: Set<number>
  at: (persistId: number) => PptRecord | null
  slideWidth: number
  slideHeight: number
}

interface ShapeContext {
  doc: DocContext
  scheme: ColorScheme
  master: MasterInfo | null
  outline: OutlineEntry[]
  /** outline entries referenced by a shape (the rest become stacked text boxes) */
  outlineUsed: Set<number>
  /** on a master: placeholders are prompts, not content */
  isMaster: boolean
  /** text boxes without any anchor stack down the slide; this counts them */
  stacked: number
}

/**
 * The presentation: slide size from the DocumentAtom, the slide order from
 * the DocumentContainer's SlideListWithText (instance 0 = slides, each
 * SlidePersistAtom followed by that slide's outline text), and each slide's
 * shapes from its drawing (plus the master's, when the slide follows it).
 */
export function parsePptPresentation(
  document: Uint8Array,
  currentUser: Uint8Array | null,
  pictureStream: Uint8Array | null = null,
): PptPresentation {
  const { directory, docPersistId } = buildPersistChain(document, currentUser)
  const chained = directory.size > 0
  const at = (persistId: number): PptRecord | null => {
    const offset = directory.get(persistId)
    return offset === undefined ? null : recordAt(document, offset)
  }

  // the live DocumentContainer: docPersistIdRef of the current UserEditAtom
  let documentRecord: PptRecord | null = null
  if (docPersistId !== null) {
    documentRecord = at(docPersistId)
    if (!documentRecord || documentRecord.type !== PPT_TYPES.documentContainer) {
      throw new LegacyPptError(
        'bad_persist',
        `docPersistIdRef ${docPersistId} is not a DocumentContainer`,
      )
    }
  } else if (!chained) {
    // a flat file (no readable change directory) holds exactly one generation
    for (const record of walkRecords(document)) {
      if (record.type === PPT_TYPES.documentContainer) {
        documentRecord = record
        break
      }
    }
  }
  if (!documentRecord) throw new LegacyPptError('corrupt', 'no DocumentContainer')

  const children = [
    ...walkRecords(document, documentRecord.offset, documentRecord.offset + documentRecord.length),
  ]

  // DocumentAtom: the slide surface, in master units
  let slideWidth = DEFAULT_SLIDE_WIDTH
  let slideHeight = DEFAULT_SLIDE_HEIGHT
  for (const record of children) {
    if (record.type !== PPT_TYPES.documentAtom || record.length < 8) continue
    const view = viewOf(payloadOf(document, record))
    const cx = view.getUint32(0, true)
    const cy = view.getUint32(4, true)
    if (cx > 0 && cy > 0 && cx <= MAX_SLIDE_MASTER_UNITS && cy <= MAX_SLIDE_MASTER_UNITS) {
      slideWidth = cx
      slideHeight = cy
    }
  }

  // Environment: fonts and the document-level text defaults; PPDrawingGroup: the picture store
  let fonts: string[] = []
  const docStyles = new Map<number, LevelStyle[]>()
  let dgg: { stream: Uint8Array; record: PptRecord } | null = null
  for (const record of children) {
    if (record.type === PPT_TYPES.environment) {
      for (const inner of walkRecordsLenient(
        document,
        record.offset,
        record.offset + record.length,
      )) {
        if (inner.type === PPT_TYPES.fontCollection) fonts = parseFontCollection(document, inner)
        else if (inner.type === PPT_TYPES.txMasterStyleAtom) {
          docStyles.set(
            inner.instance,
            parseMasterStyle(payloadOf(document, inner), inner.instance),
          )
        }
      }
    } else if (record.type === PPT_TYPES.ppDrawingGroup) {
      for (const inner of walkRecordsLenient(
        document,
        record.offset,
        record.offset + record.length,
      )) {
        if (inner.type === ESCHER.dggContainer) dgg = { stream: document, record: inner }
      }
    }
  }
  const pictures = new PictureStore(pictureStream, dgg)

  // SlideListWithText instance 1 lists the masters: each MasterPersistAtom
  // carries the persist id of a master container and the masterId slides name
  const masterPersistIds = new Map<number, number>()
  for (const record of children) {
    if (record.type !== PPT_TYPES.slideListWithText || record.instance !== 1) continue
    for (const inner of walkRecordsLenient(
      document,
      record.offset,
      record.offset + record.length,
    )) {
      if (inner.type !== PPT_TYPES.slidePersistAtom || inner.length < 16) continue
      const view = viewOf(payloadOf(document, inner))
      masterPersistIds.set(view.getUint32(12, true), view.getUint32(0, true))
    }
  }

  // SlideListWithText instance 0 is the slide list; instance 1 is the main
  // master's outline and 2 the notes', so only 0 is the deck order.
  const outlineProps9 = outlineTextProps9Of(document, children)
  const slideEntries: { persistIdRef: number; outline: OutlineEntry[] }[] = []
  for (const record of children) {
    if (record.type !== PPT_TYPES.slideListWithText || record.instance !== 0) continue
    let entry: { persistIdRef: number; outline: OutlineEntry[] } | null = null
    let props9: OutlineProps9[] = []
    for (const inner of walkRecords(document, record.offset, record.offset + record.length)) {
      if (inner.type === PPT_TYPES.slidePersistAtom) {
        if (inner.length < 4) continue
        const view = viewOf(payloadOf(document, inner))
        const persistIdRef = view.getUint32(0, true)
        entry = { persistIdRef, outline: [] }
        slideEntries.push(entry)
        props9 = inner.length >= 16 ? (outlineProps9.get(view.getUint32(12, true)) ?? []) : []
        continue
      }
      if (inner.type === PPT_TYPES.textHeaderAtom) {
        // every header counts, text or not: OutlineTextRefAtom indexes the headers
        const kind = textKindOf(document, inner)
        if (entry && kind !== null) {
          const at = props9.findIndex((p) => p.kind === kind)
          const autoNumbers = at >= 0 ? props9.splice(at, 1)[0]!.autoNumbers : undefined
          entry.outline.push({
            kind,
            raw: '',
            styles: null,
            ruler: null,
            fields: [],
            ...(autoNumbers ? { autoNumbers } : {}),
          })
        }
        continue
      }
      if (inner.type === PPT_TYPES.styleTextPropAtom) {
        const last = entry?.outline[entry.outline.length - 1]
        if (last && !last.styles && last.raw.length > 0) {
          last.styles = parseStyleRuns(payloadOf(document, inner), last.raw.length)
        }
        continue
      }
      if (inner.type === PPT_TYPES.textRulerAtom) {
        const last = entry?.outline[entry.outline.length - 1]
        if (last) last.ruler = parseTextRuler(payloadOf(document, inner))
        continue
      }
      const field = metaCharField(document, inner)
      if (field) {
        entry?.outline[entry.outline.length - 1]?.fields.push(field)
        continue
      }
      if (inner.type !== PPT_TYPES.textCharsAtom && inner.type !== PPT_TYPES.textBytesAtom) continue
      const last = entry?.outline[entry.outline.length - 1]
      if (last && last.raw.length === 0) last.raw = rawTextOf(document, inner)
    }
  }
  if (slideEntries.length > MAX_SLIDES) {
    throw new LegacyPptError('limits_exceeded', `more than ${MAX_SLIDES} slides`)
  }

  const doc: DocContext = {
    stream: document,
    fonts,
    docStyles,
    pictures,
    masters: new Map(),
    masterPersistIds,
    mastersInProgress: new Set(),
    at,
    slideWidth,
    slideHeight,
  }

  // masters: every MainMasterContainer (and title masters, which are SlideContainers
  // a SlideAtom.masterIdRef names), keyed by stream offset
  const masterRecords: PptRecord[] = []
  if (chained) {
    for (const offset of directory.values()) {
      const record = recordAt(document, offset)
      if (record?.type === PPT_TYPES.mainMasterContainer) masterRecords.push(record)
    }
  } else {
    for (const record of walkRecords(document)) {
      if (record.type === PPT_TYPES.mainMasterContainer) masterRecords.push(record)
    }
  }
  masterRecords.sort((a, b) => a.offset - b.offset)
  for (const record of masterRecords) doc.masters.set(record.offset, parseMaster(doc, record, null))
  const firstMaster = masterRecords.length
    ? (doc.masters.get(masterRecords[0]!.offset) ?? null)
    : null

  const slides: PptSlide[] = []
  const slideRecords: Array<{ record: PptRecord; outline: OutlineEntry[] }> = []
  if (chained && slideEntries.length > 0) {
    for (const entry of slideEntries) {
      const record = at(entry.persistIdRef)
      if (!record || record.type !== PPT_TYPES.slideContainer) {
        throw new LegacyPptError('bad_slide', `persist id ${entry.persistIdRef}`)
      }
      slideRecords.push({ record, outline: entry.outline })
    }
  } else {
    // no usable slide list (a flat file still knows its containers): the
    // containers in stream order are the deck order
    const containers: PptRecord[] = []
    if (chained) {
      for (const offset of directory.values()) {
        const record = recordAt(document, offset)
        if (record?.type === PPT_TYPES.slideContainer) containers.push(record)
      }
    } else {
      for (const record of walkRecords(document)) {
        if (record.type === PPT_TYPES.slideContainer) containers.push(record)
      }
    }
    containers.sort((a, b) => a.offset - b.offset)
    if (containers.length > MAX_SLIDES) {
      throw new LegacyPptError('limits_exceeded', `more than ${MAX_SLIDES} slides`)
    }
    for (const [index, record] of containers.entries()) {
      // a flat file still groups each slide's outline after its SlidePersistAtom
      const outline =
        !chained && slideEntries.length === containers.length ? slideEntries[index]!.outline : []
      slideRecords.push({ record, outline })
    }
  }
  for (const { record, outline } of slideRecords) {
    slides.push(parseSlide(doc, record, outline, firstMaster))
  }
  return { slides, slideWidth, slideHeight, pictures }
}

/* ---------- PP9 extensions ---------- */

interface OutlineProps9 {
  kind: number
  autoNumbers: Array<AutoNumber | null>
}

/** the "___PPT9" binary tag extension among these records (ProgTags → ProgBinaryTag) */
function pp9ExtensionOf(stream: Uint8Array, from: number, to: number): PptRecord | null {
  for (const record of walkRecordsLenient(stream, from, to)) {
    if (record.type === PPT_TYPES.progTags) {
      const found = pp9ExtensionOf(stream, record.offset, record.offset + record.length)
      if (found) return found
    }
    if (record.type !== PPT_TYPES.progBinaryTag) continue
    let name = ''
    for (const inner of walkRecordsLenient(stream, record.offset, record.offset + record.length)) {
      if (inner.type === PPT_TYPES.cString) name = rawTextChars(payloadOf(stream, inner))
      else if (inner.type === PPT_TYPES.binaryTagExtension && name === '___PPT9') return inner
    }
  }
  return null
}

/** a shape's StyleTextProp9Atom (in its client data), as per-paragraph-run auto-numbering */
function shapeAutoNumbersOf(
  stream: Uint8Array,
  shape: EscherShape,
): Array<AutoNumber | null> | null {
  if (!shape.clientData) return null
  const ext = pp9ExtensionOf(
    stream,
    shape.clientData.offset,
    shape.clientData.offset + shape.clientData.length,
  )
  if (!ext) return null
  for (const record of walkRecordsLenient(stream, ext.offset, ext.offset + ext.length)) {
    if (record.type === PPT_TYPES.styleTextProp9Atom) {
      return parseStyleTextProp9(payloadOf(stream, record))
    }
  }
  return null
}

/** the document's OutlineTextProps9 entries (ProgTags under the List container), by slide id, in outline order */
function outlineTextProps9Of(
  stream: Uint8Array,
  documentChildren: PptRecord[],
): Map<number, OutlineProps9[]> {
  const out = new Map<number, OutlineProps9[]>()
  for (const record of documentChildren) {
    if (record.type !== PPT_TYPES.progTags && record.type !== PPT_TYPES.list) continue
    const ext = pp9ExtensionOf(stream, record.offset, record.offset + record.length)
    if (!ext) continue
    for (const container of walkRecordsLenient(stream, ext.offset, ext.offset + ext.length)) {
      if (container.type !== PPT_TYPES.outlineTextProps9) continue
      let header: { slideId: number; kind: number } | null = null
      for (const inner of walkRecordsLenient(
        stream,
        container.offset,
        container.offset + container.length,
      )) {
        if (inner.type === PPT_TYPES.outlineTextPropsHeader9Atom && inner.length >= 8) {
          const view = viewOf(payloadOf(stream, inner))
          header = { slideId: view.getUint32(0, true), kind: view.getUint32(4, true) }
        } else if (inner.type === PPT_TYPES.styleTextProp9Atom && header) {
          const list = out.get(header.slideId) ?? []
          list.push({
            kind: header.kind,
            autoNumbers: parseStyleTextProp9(payloadOf(stream, inner)),
          })
          out.set(header.slideId, list)
          header = null
        }
      }
    }
  }
  return out
}

/* ---------- containers ---------- */

function slideAtomOf(stream: Uint8Array, container: PptRecord): SlideAtomInfo | null {
  for (const record of walkRecordsLenient(
    stream,
    container.offset,
    container.offset + container.length,
  )) {
    if (record.type !== PPT_TYPES.slideAtom || record.length < 22) continue
    const view = viewOf(payloadOf(stream, record))
    const flags = view.getUint16(20, true)
    return {
      masterIdRef: view.getUint32(12, true),
      masterObjects: (flags & 0x1) !== 0,
      masterScheme: (flags & 0x2) !== 0,
      masterBackground: (flags & 0x4) !== 0,
    }
  }
  return null
}

function schemeOf(stream: Uint8Array, container: PptRecord): ColorScheme | null {
  for (const record of walkRecordsLenient(
    stream,
    container.offset,
    container.offset + container.length,
  )) {
    if (record.type !== PPT_TYPES.colorSchemeAtom || record.instance !== 1 || record.length < 32)
      continue
    const view = viewOf(payloadOf(stream, record))
    const colors: string[] = []
    for (let index = 0; index < 8; index++) {
      colors.push(
        officeArtColor(view.getUint32(index * 4, true) & 0xffffff, DEFAULT_SCHEME) ?? '000000',
      )
    }
    return { colors }
  }
  return null
}

function drawingOf(stream: Uint8Array, container: PptRecord): EscherShape[] {
  for (const record of walkRecordsLenient(
    stream,
    container.offset,
    container.offset + container.length,
  )) {
    if (record.type !== PPT_TYPES.ppDrawing) continue
    for (const inner of walkRecordsLenient(stream, record.offset, record.offset + record.length)) {
      if (inner.type === ESCHER.dgContainer) return parseDrawing(stream, inner)
    }
  }
  return []
}

function placeholderType(stream: Uint8Array, shape: EscherShape): number | null {
  if (!shape.clientData) return null
  for (const record of walkRecordsLenient(
    stream,
    shape.clientData.offset,
    shape.clientData.offset + shape.clientData.length,
  )) {
    if (record.type === PPT_TYPES.oePlaceholderAtom && record.length >= 8) {
      return stream[record.offset + 4]!
    }
  }
  return null
}

/**
 * The master placeholder a slide placeholder inherits from (PlaceholderEnum:
 * 13 title, 14 body, 15 centre title, 16 subtitle, 17/18 vertical title/body,
 * 19-24 object/graph/table/clip-art/org-chart/media (body-shaped),
 * 25 slide number, 26 footer, 27 header, 28 date → the master's 1-10).
 */
function masterPlaceholderType(slideType: number): number {
  switch (slideType) {
    case 13:
    case 17:
      return 1
    case 15:
      return 3
    case 14:
    case 18:
    case 19:
    case 20:
    case 21:
    case 22:
    case 23:
    case 24:
      return 2
    case 16:
      return 4
    case 28:
      return 7
    case 25:
      return 8
    case 26:
      return 9
    case 27:
      return 10
    default:
      return slideType
  }
}

function parseMaster(doc: DocContext, record: PptRecord, parent: MasterInfo | null): MasterInfo {
  const stream = doc.stream
  const atom = slideAtomOf(stream, record)
  const shapes = drawingOf(stream, record)
  const styles = new Map<number, LevelStyle[]>(parent?.styles ?? [])
  for (const inner of walkRecordsLenient(stream, record.offset, record.offset + record.length)) {
    if (inner.type === PPT_TYPES.txMasterStyleAtom) {
      styles.set(inner.instance, parseMasterStyle(payloadOf(stream, inner), inner.instance))
    }
  }
  const scheme = schemeOf(stream, record) ?? parent?.scheme ?? null
  const placeholders = new Map<number, EscherShape>(parent?.placeholders ?? [])
  for (const shape of shapes) {
    const type = placeholderType(stream, shape)
    if (type === null) continue
    placeholders.set(type, shape)
  }
  // a centre title / subtitle falls back to the title / body master placeholder
  if (!placeholders.has(3) && placeholders.has(1)) placeholders.set(3, placeholders.get(1)!)
  if (!placeholders.has(4) && placeholders.has(2)) placeholders.set(4, placeholders.get(2)!)
  if (!placeholders.has(1) && placeholders.has(3)) placeholders.set(1, placeholders.get(3)!)
  const info: MasterInfo = { shapes, scheme, styles, placeholders, atom }
  const background = backgroundOf(stream, shapes, scheme ?? DEFAULT_SCHEME, doc.pictures)
  info.background = background ?? parent?.background
  return info
}

/** the master a SlideAtom names; a title master (a SlideContainer) is parsed on first use */
function masterFor(
  doc: DocContext,
  atom: SlideAtomInfo | null,
  fallback: MasterInfo | null,
): MasterInfo | null {
  if (!atom || atom.masterIdRef === 0) return fallback
  const persistId = doc.masterPersistIds.get(atom.masterIdRef)
  const record = persistId === undefined ? null : doc.at(persistId)
  if (!record) return fallback
  const cached = doc.masters.get(record.offset)
  if (cached) return cached
  if (record.type === PPT_TYPES.slideContainer) {
    if (doc.mastersInProgress.has(atom.masterIdRef)) return fallback
    doc.mastersInProgress.add(atom.masterIdRef)
    try {
      const own = slideAtomOf(doc.stream, record)
      const parent =
        own && own.masterIdRef !== atom.masterIdRef ? masterFor(doc, own, fallback) : fallback
      const info = parseMaster(doc, record, parent)
      doc.masters.set(record.offset, info)
      return info
    } finally {
      doc.mastersInProgress.delete(atom.masterIdRef)
    }
  }
  return fallback
}

function parseSlide(
  doc: DocContext,
  slideRecord: PptRecord,
  outline: readonly OutlineEntry[],
  firstMaster: MasterInfo | null,
): PptSlide {
  const stream = doc.stream
  const atom = slideAtomOf(stream, slideRecord)
  const master = masterFor(doc, atom, firstMaster)
  const ownScheme = schemeOf(stream, slideRecord)
  const scheme =
    (atom?.masterScheme ? (master?.scheme ?? ownScheme) : (ownScheme ?? master?.scheme)) ??
    DEFAULT_SCHEME
  const escherShapes = drawingOf(stream, slideRecord)

  const ctx: ShapeContext = {
    doc,
    scheme,
    master,
    outline: [...outline],
    outlineUsed: new Set(),
    isMaster: false,
    stacked: 0,
  }
  const shapes: PptAnyShape[] = []
  if (master && (atom?.masterObjects ?? true)) {
    const masterCtx: ShapeContext = {
      ...ctx,
      outline: [],
      outlineUsed: new Set(),
      isMaster: true,
      stacked: 0,
    }
    for (const shape of master.shapes) collectShapes(masterCtx, shape, null, 0, shapes)
  }
  for (const shape of escherShapes) collectShapes(ctx, shape, null, 0, shapes)

  const ownBackground = backgroundOf(stream, escherShapes, scheme, doc.pictures)
  const background = (atom?.masterBackground
    ? (master?.background ?? ownBackground)
    : (ownBackground ?? master?.background)) ?? {
    kind: 'solid' as const,
    color: scheme.colors[0] ?? 'FFFFFF',
  }

  // outline text no shape claimed (a placeholder whose drawing is missing) still lands on the slide
  const textSequences: PptSlideText[] = []
  const push = (sequence: PptSlideText): void => {
    if (sequence.text.trim().length === 0) return
    if (textSequences.length >= MAX_TEXT_SEQUENCES) {
      throw new LegacyPptError(
        'limits_exceeded',
        `more than ${MAX_TEXT_SEQUENCES} text sequences on one slide`,
      )
    }
    textSequences.push(sequence)
  }
  for (const entry of outline) push({ kind: entry.kind, text: normalizeRaw(entry.raw) })
  for (const sequence of textBoxesOf(stream, slideRecord)) push(sequence)
  outline.forEach((entry, index) => {
    if (ctx.outlineUsed.has(index) || entry.raw.trim().length === 0) return
    const body = textBodyFor(
      ctx,
      entry.kind,
      entry.raw,
      entry.styles,
      new Map(),
      entry.fields,
      entry.ruler,
      undefined,
      entry.autoNumbers ?? null,
    )
    const anchor = master?.placeholders.get(
      masterPlaceholderType(entry.kind === 0 ? 13 : 14),
    )?.anchor
    const rect = anchor ?? stackedRect(ctx)
    shapes.push({
      type: 'shape',
      ...rectToEmu(rect),
      rotation: 0,
      flipH: false,
      flipV: false,
      text: body,
    })
  })
  return { textSequences, shapes, background }
}

/* ---------- shapes ---------- */

const BOOLEAN_PIDS = new Set<number>([
  OPT.textBooleans,
  OPT.fillBooleans,
  OPT.lineBooleans,
  OPT.shadowBooleans,
  OPT.groupShapeBooleans,
  0x033f,
])

/**
 * A slide shape's properties over its master placeholder's. A boolean set is
 * merged bit by bit: the slide only speaks for the bits whose use flag it
 * sets, the master keeps the rest.
 */
function mergeProps(base: Map<number, number>, own: Map<number, number>): Map<number, number> {
  const out = new Map(base)
  for (const [pid, value] of own) {
    const inherited = out.get(pid)
    if (inherited !== undefined && BOOLEAN_PIDS.has(pid)) {
      const useMask = (value >>> 16) & 0xffff
      const merged = (inherited & ~useMask & ~(useMask << 16)) | (value & useMask) | (useMask << 16)
      out.set(pid, merged >>> 0)
    } else out.set(pid, value)
  }
  return out
}

/** the next slot of the stage-1 layout: a text box with no anchor lands in a row of its own */
function stackedRect(ctx: ShapeContext): Rect {
  const margin = 288 // 0.5 in
  const rowHeight = Math.max(380, Math.floor((ctx.doc.slideHeight - margin * 2) / 6))
  const top = margin + ctx.stacked * rowHeight
  ctx.stacked++
  return { left: margin, top, right: ctx.doc.slideWidth - margin, bottom: top + rowHeight }
}

function rectToEmu(rect: Rect): { x: number; y: number; cx: number; cy: number } {
  const left = Math.min(rect.left, rect.right)
  const top = Math.min(rect.top, rect.bottom)
  return {
    x: Math.round(left * EMU_PER_MASTER_UNIT),
    y: Math.round(top * EMU_PER_MASTER_UNIT),
    cx: Math.round(Math.abs(rect.right - rect.left) * EMU_PER_MASTER_UNIT),
    cy: Math.round(Math.abs(rect.bottom - rect.top) * EMU_PER_MASTER_UNIT),
  }
}

function mapChildRect(child: Rect, bounds: Rect, outer: Rect): Rect | null {
  const boundsWidth = bounds.right - bounds.left
  const boundsHeight = bounds.bottom - bounds.top
  if (boundsWidth === 0 || boundsHeight === 0) return null
  const scaleX = (outer.right - outer.left) / boundsWidth
  const scaleY = (outer.bottom - outer.top) / boundsHeight
  return {
    left: outer.left + (child.left - bounds.left) * scaleX,
    top: outer.top + (child.top - bounds.top) * scaleY,
    right: outer.left + (child.right - bounds.left) * scaleX,
    bottom: outer.top + (child.bottom - bounds.top) * scaleY,
  }
}

/** the rotation and flips a group hands down to its members */
interface GroupTransform {
  /** degrees clockwise, about the group's centre */
  rotation: number
  flipH: boolean
  flipV: boolean
  centerX: number
  centerY: number
}

function shapeRotation(shape: EscherShape): number {
  const raw = shape.props.get(OPT.rotation)
  return raw === undefined ? 0 : ((fixedProp(raw) % 360) + 360) % 360
}

/**
 * A member rectangle in slide space after the group's own transform: its
 * centre mirrored and rotated about the group's centre (the size is kept;
 * the member's own rotation and flips are adjusted in buildShape).
 */
function applyGroupTransform(rect: Rect, transform: GroupTransform): Rect {
  let cx = (rect.left + rect.right) / 2
  let cy = (rect.top + rect.bottom) / 2
  if (transform.flipH) cx = 2 * transform.centerX - cx
  if (transform.flipV) cy = 2 * transform.centerY - cy
  if (transform.rotation) {
    const rad = (transform.rotation * Math.PI) / 180
    const dx = cx - transform.centerX
    const dy = cy - transform.centerY
    cx = transform.centerX + dx * Math.cos(rad) - dy * Math.sin(rad)
    cy = transform.centerY + dx * Math.sin(rad) + dy * Math.cos(rad)
  }
  const halfW = Math.abs(rect.right - rect.left) / 2
  const halfH = Math.abs(rect.bottom - rect.top) / 2
  return { left: cx - halfW, top: cy - halfH, right: cx + halfW, bottom: cy + halfH }
}

function isQuarterTurned(rotation: number): boolean {
  return (rotation > 45 && rotation < 135) || (rotation > 225 && rotation < 315)
}

/** the same box with its sides swapped about the centre */
function swapSides(rect: Rect): Rect {
  const cx = (rect.left + rect.right) / 2
  const cy = (rect.top + rect.bottom) / 2
  const halfW = Math.abs(rect.right - rect.left) / 2
  const halfH = Math.abs(rect.bottom - rect.top) / 2
  return { left: cx - halfH, top: cy - halfW, right: cx + halfH, bottom: cy + halfW }
}

/** flatten the shape tree into slide-space shapes (groups carry no paint of their own) */
function collectShapes(
  ctx: ShapeContext,
  shape: EscherShape,
  rect: Rect | null,
  depth: number,
  out: PptAnyShape[],
  inherited: GroupTransform | null = null,
): void {
  if (depth > MAX_GROUP_DEPTH || out.length >= MAX_SHAPES_PER_SLIDE) return
  if (shape.flags & FSP.deleted) return
  if (shape.flags & FSP.background) return
  if (boolProp(shape.props, OPT.groupShapeBooleans, 0x02) === true) return
  let own = rect ?? shape.anchor ?? null
  if (shape.flags & FSP.group) {
    if (!own || !shape.groupBounds) return
    const rotation = shapeRotation(shape)
    // a group's anchor is its rotated bounding box too (see buildShape)
    if (isQuarterTurned(rotation)) own = swapSides(own)
    const flipH = (shape.flags & FSP.flipH) !== 0
    const flipV = (shape.flags & FSP.flipV) !== 0
    // `own` already sits where the groups above put it, so what the members
    // still need is the net rotation/flips, about this group's centre
    const ownTransform: GroupTransform | null =
      rotation || flipH || flipV
        ? {
            rotation,
            flipH,
            flipV,
            centerX: (own.left + own.right) / 2,
            centerY: (own.top + own.bottom) / 2,
          }
        : null
    const combined = combineTransforms(inherited, ownTransform, own)
    for (const child of shape.children) {
      if (!child.childAnchor) continue
      let childRect = mapChildRect(child.childAnchor, shape.groupBounds, own)
      if (!childRect) continue
      if (combined) childRect = applyGroupTransform(childRect, combined)
      collectShapes(ctx, child, childRect, depth + 1, out, combined)
    }
    return
  }
  const built = buildShape(ctx, shape, own, inherited)
  if (built) out.push(built)
}

/** the net rotation/flips a nested member receives, re-centred on the inner group */
function combineTransforms(
  outer: GroupTransform | null,
  inner: GroupTransform | null,
  innerRect: Rect,
): GroupTransform | null {
  if (!outer) return inner
  const centerX = (innerRect.left + innerRect.right) / 2
  const centerY = (innerRect.top + innerRect.bottom) / 2
  if (!inner) return { ...outer, centerX, centerY }
  // a mirrored group turns its rotation the other way
  const innerRotation = outer.flipH !== outer.flipV ? -inner.rotation : inner.rotation
  return {
    rotation: (((outer.rotation + innerRotation) % 360) + 360) % 360,
    flipH: outer.flipH !== inner.flipH,
    flipV: outer.flipV !== inner.flipV,
    centerX,
    centerY,
  }
}

function buildShape(
  ctx: ShapeContext,
  raw: EscherShape,
  rect: Rect | null,
  group: GroupTransform | null = null,
): PptAnyShape | null {
  const stream = ctx.doc.stream
  const placeholder = placeholderType(stream, raw)
  // a slide placeholder inherits the master placeholder's properties under its
  // own; so does a title/body text shape without a placeholder atom (PowerPoint
  // matches it to the master placeholder by text type), except the text anchor
  let masterType = placeholder !== null ? masterPlaceholderType(placeholder) : null
  if (masterType === null && !ctx.isMaster) {
    const kind = textKindOfShape(ctx, raw)
    if (kind === 0 || kind === 6) masterType = 1
    else if (kind === 1 || kind === 5 || kind === 7 || kind === 8) masterType = 2
  }
  const inherited =
    masterType !== null && !ctx.isMaster ? ctx.master?.placeholders.get(masterType) : undefined
  let shape = raw
  if (inherited) {
    const base = new Map(inherited.props)
    if (placeholder === null) base.delete(OPT.anchorText)
    shape = {
      ...raw,
      props: mergeProps(base, raw.props),
      complex: new Map([...inherited.complex, ...raw.complex]),
    }
  }
  let text = textBodyOf(ctx, shape)
  // a date / slide-number placeholder whose only content is a bare '*' with
  // no MetaCharacter atom behind it has nothing to show
  if (text && placeholder !== null && placeholder >= 25 && placeholder <= 28) {
    const runs = text.paragraphs.flatMap((p) => p.runs)
    const plain = runs
      .map((r) => r.text)
      .join('')
      .trim()
    if (plain === '*' && !runs.some((r) => r.field)) text = undefined
  }
  const pib = shape.props.get(OPT.pib)
  const picture = pib !== undefined ? ctx.doc.pictures.get(pib) : null

  let anchor = rect ?? inherited?.anchor ?? null
  if (!anchor && text && !ctx.isMaster) anchor = stackedRect(ctx)
  if (!anchor) return null
  // a placeholder on a master is a prompt; on a slide an empty one shows nothing
  if (placeholder !== null && (ctx.isMaster || (!text && !picture))) return null
  // master decorations the slide hides: date/footer/number placeholders are handled above
  const geometry = rectToEmu(anchor)
  const ownRotation = shapeRotation(shape)
  const ownFlipH = (shape.flags & FSP.flipH) !== 0
  const ownFlipV = (shape.flags & FSP.flipV) !== 0
  // inside a mirrored group the member's own rotation runs the other way
  const mirrored = group ? group.flipH !== group.flipV : false
  const rotation = group
    ? (((group.rotation + (mirrored ? -ownRotation : ownRotation)) % 360) + 360) % 360
    : ownRotation
  // Escher keeps the rotated bounding box for a shape turned near 90° or
  // 270° (by its own rotation; a group's turn is applied on top); DrawingML
  // wants the unrotated one, so the sides swap about the centre
  if (isQuarterTurned(ownRotation)) {
    const centerX = geometry.x + geometry.cx / 2
    const centerY = geometry.y + geometry.cy / 2
    ;[geometry.cx, geometry.cy] = [geometry.cy, geometry.cx]
    geometry.x = Math.round(centerX - geometry.cx / 2)
    geometry.y = Math.round(centerY - geometry.cy / 2)
  }
  const flipH = group ? ownFlipH !== group.flipH : ownFlipH
  const flipV = group ? ownFlipV !== group.flipV : ownFlipV
  const adjust =
    shape.shapeType === 19 ? placeArc(shape, geometry, rotation, flipH, flipV) : undefined
  const base = { ...geometry, rotation, flipH, flipV }
  const line = lineOf(ctx, shape)
  // a pib the store cannot decode still lands as a picture entry so the writer can count it
  if (pib !== undefined && !text && (picture || shape.shapeType === 75)) {
    const crop = cropOf(shape)
    const transparent = transparentColorOf(ctx, shape, picture)
    return {
      type: 'picture',
      ...base,
      pib,
      line,
      ...(crop ? { crop } : {}),
      ...(transparent ? { transparent } : {}),
    }
  }
  const fill = fillOf(ctx, shape, raw, placeholder !== null)
  let preset = PRESET_BY_SHAPE_TYPE[shape.shapeType]
  let path: CustomPath | undefined
  if (shape.shapeType === 0 || preset === undefined) {
    path = freeformPath(shape) ?? undefined
    if (!path) preset = 'rect'
  }
  if (!fill && !line && !text) return null
  return { type: 'shape', ...base, preset, path, adjust, fill, line, text }
}

/**
 * An arc's anchor frames only the swept sector, DrawingML's `arc` the whole ellipse:
 * grow the geometry (the anchor maps the geometry box), keep the anchor's centre of
 * rotation, and return the start/end angles clockwise from 3 o'clock as `arc` reads them.
 */
function placeArc(
  shape: EscherShape,
  geometry: { x: number; y: number; cx: number; cy: number },
  rotation: number,
  flipH: boolean,
  flipV: boolean,
): Record<string, number> | undefined {
  const { box, ellipse, start, end } = arcGeometry(shape)
  const boxW = box.right - box.left
  const boxH = box.bottom - box.top
  if (boxW <= 0 || boxH <= 0) return undefined
  const sx = geometry.cx / boxW
  const sy = geometry.cy / boxH
  const mapX = (g: number): number =>
    flipH ? geometry.x + geometry.cx - (g - box.left) * sx : geometry.x + (g - box.left) * sx
  const mapY = (g: number): number =>
    flipV ? geometry.y + geometry.cy - (g - box.top) * sy : geometry.y + (g - box.top) * sy
  const left = Math.min(mapX(ellipse.left), mapX(ellipse.right))
  const top = Math.min(mapY(ellipse.top), mapY(ellipse.bottom))
  const cx = Math.abs(ellipse.right - ellipse.left) * sx
  const cy = Math.abs(ellipse.bottom - ellipse.top) * sy
  const anchorCenterX = geometry.x + geometry.cx / 2
  const anchorCenterY = geometry.y + geometry.cy / 2
  let centerX = left + cx / 2
  let centerY = top + cy / 2
  if (rotation) {
    const a = (rotation * Math.PI) / 180
    const dx = centerX - anchorCenterX
    const dy = centerY - anchorCenterY
    centerX = anchorCenterX + dx * Math.cos(a) - dy * Math.sin(a)
    centerY = anchorCenterY + dx * Math.sin(a) + dy * Math.cos(a)
  }
  geometry.x = Math.round(centerX - cx / 2)
  geometry.y = Math.round(centerY - cy / 2)
  geometry.cx = Math.round(cx)
  geometry.cy = Math.round(cy)
  const ellipseCenterX = (ellipse.left + ellipse.right) / 2
  const ellipseCenterY = (ellipse.top + ellipse.bottom) / 2
  const angleOf = ([x, y]: [number, number]): number => {
    const degrees =
      (Math.atan2(
        (y - ellipseCenterY) / (ellipse.bottom - ellipse.top),
        (x - ellipseCenterX) / (ellipse.right - ellipse.left),
      ) *
        180) /
      Math.PI
    return ((degrees % 360) + 360) % 360
  }
  return { adj1: Math.round(angleOf(start) * 60000), adj2: Math.round(angleOf(end) * 60000) }
}

/** the colour a raster picture shows through; metafiles carry their own transparency */
function transparentColorOf(
  ctx: ShapeContext,
  shape: EscherShape,
  picture: PictureMedia | null,
): string | undefined {
  const value = shape.props.get(OPT.pictureTransparent)
  if (value === undefined || !picture) return undefined
  if (picture.mediaType === 'image/x-emf' || picture.mediaType === 'image/x-wmf') return undefined
  return officeArtColor(value, ctx.scheme, {
    fill: shape.props.get(OPT.fillColor),
    back: shape.props.get(OPT.fillBackColor),
    line: shape.props.get(OPT.lineColor),
  })
}

function cropOf(shape: EscherShape): PptPicture['crop'] | undefined {
  const fraction = (pid: number): number => {
    const value = shape.props.get(pid)
    if (value === undefined) return 0
    const signed = value | 0
    if (signed <= -0x10000 || signed >= 0x10000) return 0
    return Math.round((signed * 100000) / 0x10000)
  }
  const crop = {
    l: fraction(OPT.cropFromLeft),
    t: fraction(OPT.cropFromTop),
    r: fraction(OPT.cropFromRight),
    b: fraction(OPT.cropFromBottom),
  }
  if (crop.l + crop.r >= 100000 || crop.t + crop.b >= 100000) return undefined
  return crop.l || crop.t || crop.r || crop.b ? crop : undefined
}

const LINE_TYPES = new Set([20, 32, 33, 34, 35, 36, 37, 38, 39, 40])

/**
 * Whether a shape paints a fill: its own fFilled when it says so, else a fill type other
 * than the plain solid default or a literal (non-scheme) fill colour of its own opts in (a
 * slide title given its own shade is filled even though the master's title box is not;
 * PowerPoint writes fillType solid, the opacities and the scheme colours on shapes it saves
 * without meaning a fill), else the inherited fFilled. With none of those a placeholder is unfilled — PowerPoint
 * writes the scheme fill/back colours on every placeholder without meaning a fill — and
 * any other shape takes its type's default (a rectangle is filled, a text box is not).
 */
function fillOf(
  ctx: ShapeContext,
  shape: EscherShape,
  own: EscherShape,
  isPlaceholder: boolean,
): PptFill | undefined {
  const type = shape.shapeType
  const defaultFilled = !(
    type === 202 ||
    type === 75 ||
    LINE_TYPES.has(type) ||
    (type === 0 && !shape.complex.has(OPT.pVertices))
  )
  const ownColor = own.props.get(OPT.fillColor)
  const optedIn =
    (own.props.get(OPT.fillType) ?? 0) !== 0 ||
    (ownColor !== undefined && ((ownColor >>> 24) & 0x19) === 0)
  const filled =
    boolProp(own.props, OPT.fillBooleans, 0x10) ??
    (optedIn
      ? true
      : (boolProp(shape.props, OPT.fillBooleans, 0x10) ?? (isPlaceholder ? false : defaultFilled)))
  if (!filled) return undefined
  return fillFromProps(shape, ctx.scheme, ctx.doc.pictures)
}

function fillFromProps(
  shape: EscherShape,
  scheme: ColorScheme,
  pictures: PictureStore,
): PptFill | undefined {
  const fillType = shape.props.get(OPT.fillType) ?? 0
  const context = {
    fill: shape.props.get(OPT.fillColor),
    back: shape.props.get(OPT.fillBackColor),
    line: shape.props.get(OPT.lineColor),
  }
  const color =
    officeArtColor(shape.props.get(OPT.fillColor) ?? 0xffffff, scheme, context) ?? 'FFFFFF'
  const opacity = shape.props.get(OPT.fillOpacity)
  const alpha = opacity === undefined ? undefined : Math.max(0, Math.min(1, fixedProp(opacity)))
  if (fillType === 1) {
    const pib = shape.props.get(OPT.fillBlip)
    const media = pib === undefined ? null : pictures.get(pib)
    const mask = media ? patternMaskOf(media) : null
    if (mask) {
      const preset = patternPreset(mask)
      const schemeIndexed = (value: number | undefined): boolean =>
        value !== undefined && ((value >>> 24) & 0x08) !== 0
      const palette =
        schemeIndexed(context.fill) || schemeIndexed(context.back) ? patternPaletteOf(media!) : null
      if (palette) return { kind: 'pattern', preset, ...palette }
      const bg =
        officeArtColor(shape.props.get(OPT.fillBackColor) ?? 0xffffff, scheme, context) ?? 'FFFFFF'
      return { kind: 'pattern', preset, fg: color, bg }
    }
    return { kind: 'solid', color, alpha }
  }
  if (fillType === 2 || fillType === 3) {
    const pib = shape.props.get(OPT.fillBlip)
    if (pib !== undefined && pictures.get(pib)) {
      return fillType === 2 ? { kind: 'picture', pib, tile: true } : { kind: 'picture', pib }
    }
    return { kind: 'solid', color, alpha }
  }
  if (fillType >= 4 && fillType <= 8) {
    const back =
      officeArtColor(shape.props.get(OPT.fillBackColor) ?? 0xffffff, scheme, context) ?? 'FFFFFF'
    let stops: Array<{ pos: number; color: string }> = []
    const shadeColors = shape.complex.get(OPT.fillShadeColors)
    if (shadeColors && shadeColors.length >= 6) {
      const view = viewOf(shadeColors)
      const count = view.getUint16(0, true)
      for (let index = 0; index < count && 6 + index * 8 + 8 <= shadeColors.length; index++) {
        const at = 6 + index * 8
        const stopColor = officeArtColor(view.getUint32(at, true), scheme, context)
        const pos = Math.max(0, Math.min(100, fixedProp(view.getUint32(at + 4, true)) * 100))
        if (stopColor) stops.push({ pos, color: stopColor })
      }
      stops.sort((a, b) => a.pos - b.pos)
    }
    const angleRaw = shape.props.get(OPT.fillAngle)
    const escherAngle = angleRaw === undefined ? 0 : fixedProp(angleRaw)
    if (stops.length < 2) {
      stops = focusStops(shape.props.get(OPT.fillFocus) ?? 0, escherAngle, color, back)
    }
    // Escher angles run counter-clockwise from "fill colour on top"; DrawingML's run clockwise from 3 o'clock
    const angle = (((90 - escherAngle) % 360) + 360) % 360
    if (fillType === 5 || fillType === 6) {
      const to = (pid: number): number => {
        const raw = shape.props.get(pid)
        return raw === undefined ? 0 : Math.max(0, Math.min(1, fixedProp(raw)))
      }
      // Escher places the focus rectangle's right/bottom edges from the left/top; DrawingML insets them
      const radial = {
        path: fillType === 6 ? ('shape' as const) : ('rect' as const),
        l: to(OPT.fillToLeft),
        t: to(OPT.fillToTop),
        r: 1 - to(OPT.fillToRight),
        b: 1 - to(OPT.fillToBottom),
      }
      return { kind: 'gradient', stops, angle, radial }
    }
    return { kind: 'gradient', stops, angle }
  }
  if (fillType === 9) return { kind: 'background' }
  return { kind: 'solid', color, alpha }
}

/**
 * Two-colour shade stops from fillFocus: 100 runs fill → back, 0 runs back → fill and
 * anything between mirrors with the fill colour at both ends and the back colour at
 * |focus| (the sign changes nothing on screen). An angle wrapped past ±180° (PowerPoint
 * shows such a mirrored shade with the colours exchanged) swaps the pair.
 */
function focusStops(
  focusRaw: number,
  angle: number,
  fill: string,
  back: string,
): Array<{ pos: number; color: string }> {
  const focus = Math.abs(Math.max(-100, Math.min(100, focusRaw | 0)))
  if (focus === 100)
    return [
      { pos: 0, color: fill },
      { pos: 100, color: back },
    ]
  if (focus === 0)
    return [
      { pos: 0, color: back },
      { pos: 100, color: fill },
    ]
  const [ends, middle] = Math.abs(angle) > 180 ? [back, fill] : [fill, back]
  return [
    { pos: 0, color: ends },
    { pos: focus, color: middle },
    { pos: 100, color: ends },
  ]
}

function lineOf(ctx: ShapeContext, shape: EscherShape): PptLine | undefined {
  const type = shape.shapeType
  const defaultLine = !(type === 202 || type === 75 || type === 0)
  if (!(boolProp(shape.props, OPT.lineBooleans, 0x08) ?? defaultLine)) return undefined
  const color =
    officeArtColor(shape.props.get(OPT.lineColor) ?? 0x000000, ctx.scheme, {
      fill: shape.props.get(OPT.fillColor),
      back: shape.props.get(OPT.fillBackColor),
      line: shape.props.get(OPT.lineColor),
    }) ?? '000000'
  const dash = LINE_DASHES[shape.props.get(OPT.lineDashing) ?? 0]
  const compound = LINE_COMPOUNDS[shape.props.get(OPT.lineStyle) ?? 0]
  return {
    color,
    width: shape.props.get(OPT.lineWidth) ?? 9525,
    ...(dash ? { dash } : {}),
    ...(compound ? { compound } : {}),
  }
}

/** MSOLINEDASHING → ST_PresetLineDashVal (solid omitted) */
const LINE_DASHES: Array<string | undefined> = [
  undefined,
  'sysDash',
  'sysDot',
  'sysDashDot',
  'sysDashDotDot',
  'dot',
  'dash',
  'lgDash',
  'dashDot',
  'lgDashDot',
  'lgDashDotDot',
]

/** MSOLINESTYLE → ST_CompoundLine (single omitted) */
const LINE_COMPOUNDS: Array<string | undefined> = [
  undefined,
  'dbl',
  'thickThin',
  'thinThick',
  'tri',
]

/** the master-level background: the drawing's background shape, or nothing */
function backgroundOf(
  stream: Uint8Array,
  shapes: EscherShape[],
  scheme: ColorScheme,
  pictures: PictureStore,
): PptFill | undefined {
  void stream
  for (const shape of shapes) {
    if (!(shape.flags & FSP.background)) continue
    const filled = boolProp(shape.props, OPT.fillBooleans, 0x10) ?? true
    if (!filled) return undefined
    const fill = fillFromProps(shape, scheme, pictures)
    return fill?.kind === 'background' ? undefined : fill
  }
  return undefined
}

/* ---------- text ---------- */

/** the TextHeaderEnum of a shape's text (through its outline reference when needed) */
function textKindOfShape(ctx: ShapeContext, shape: EscherShape): number | null {
  if (!shape.textbox) return null
  const stream = ctx.doc.stream
  for (const inner of walkRecordsLenient(
    stream,
    shape.textbox.offset,
    shape.textbox.offset + shape.textbox.length,
  )) {
    if (inner.type === PPT_TYPES.textHeaderAtom) return textKindOf(stream, inner)
    if (inner.type === PPT_TYPES.outlineTextRefAtom && inner.length >= 4) {
      return ctx.outline[viewOf(payloadOf(stream, inner)).getInt32(0, true)]?.kind ?? null
    }
  }
  return null
}

function textBodyOf(ctx: ShapeContext, shape: EscherShape): PptTextBody | undefined {
  if (!shape.textbox) return undefined
  const stream = ctx.doc.stream
  let kind = 4
  let raw: string | null = null
  let styles: StyleRuns | null = null
  let ruler: TextRuler | null = null
  let outlineIndex: number | null = null
  let autoNumbers = shapeAutoNumbersOf(stream, shape)
  let fields: TextField[] = []
  for (const inner of walkRecordsLenient(
    stream,
    shape.textbox.offset,
    shape.textbox.offset + shape.textbox.length,
  )) {
    if (inner.type === PPT_TYPES.textHeaderAtom) kind = textKindOf(stream, inner) ?? kind
    else if (inner.type === PPT_TYPES.textCharsAtom || inner.type === PPT_TYPES.textBytesAtom) {
      raw = rawTextOf(stream, inner)
    } else if (inner.type === PPT_TYPES.styleTextPropAtom && raw !== null) {
      styles = parseStyleRuns(payloadOf(stream, inner), raw.length)
    } else if (inner.type === PPT_TYPES.textRulerAtom) {
      ruler = parseTextRuler(payloadOf(stream, inner))
    } else if (inner.type === PPT_TYPES.outlineTextRefAtom && inner.length >= 4) {
      outlineIndex = viewOf(payloadOf(stream, inner)).getInt32(0, true)
    } else {
      const field = metaCharField(stream, inner)
      if (field) fields.push(field)
    }
  }
  if (raw === null && outlineIndex !== null) {
    const entry = ctx.outline[outlineIndex]
    if (entry) {
      ctx.outlineUsed.add(outlineIndex)
      kind = entry.kind
      raw = entry.raw
      styles = entry.styles
      ruler ??= entry.ruler
      fields = entry.fields
      autoNumbers ??= entry.autoNumbers ?? null
    }
  }
  if (raw === null || raw.trim().length === 0) return undefined
  return textBodyFor(
    ctx,
    kind,
    raw,
    styles,
    shape.props,
    fields,
    ruler,
    shape.complex.get(OPT.metroBlob),
    autoNumbers,
  )
}

function textBodyFor(
  ctx: ShapeContext,
  kind: number,
  raw: string,
  styles: StyleRuns | null,
  props: Map<number, number>,
  fields: TextField[] = [],
  ruler: TextRuler | null = null,
  metroBlob?: Uint8Array,
  autoNumbers: Array<AutoNumber | null> | null = null,
): PptTextBody {
  const { levels, base } = levelsForText(
    kind,
    ctx.isMaster,
    (instance) => ctx.master?.styles.get(instance),
    (instance) => ctx.doc.docStyles.get(instance),
  )
  const paragraphs = resolveText(raw, styles, {
    levels,
    base,
    fonts: ctx.doc.fonts,
    scheme: ctx.scheme,
    ...(ruler ? { ruler } : {}),
    ...(autoNumbers ? { autoNumbers } : {}),
  })
  if (metroBlob) applyMetroIndents(paragraphs, metroBlob)
  markFields(paragraphs, fields)
  const anchorRaw = props.get(OPT.anchorText) ?? 0
  // msoanchor: top, middle, bottom, then the centred three, then the
  // top/bottom baseline pair and their centred twins
  const anchorMap: Array<'t' | 'ctr' | 'b'> = ['t', 'ctr', 'b', 't', 'ctr', 'b', 't', 'b', 't', 'b']
  const anchorCentered = (anchorRaw >= 3 && anchorRaw <= 5) || anchorRaw === 8 || anchorRaw === 9
  const wrapRaw = props.get(OPT.wrapText) ?? 0
  const flow = props.get(OPT.textFlow) ?? 0
  return {
    kind,
    paragraphs,
    insets: [
      props.get(OPT.dxTextLeft) ?? 91440,
      props.get(OPT.dyTextTop) ?? 45720,
      props.get(OPT.dxTextRight) ?? 91440,
      props.get(OPT.dyTextBottom) ?? 45720,
    ],
    anchor: anchorMap[anchorRaw] ?? 't',
    anchorCenter: anchorCentered,
    wrap: wrapRaw !== 2,
    autofit: boolProp(props, OPT.textBooleans, 0x02) ?? false,
    vertical: flow === 1 || flow === 3 ? 'vert' : flow === 5 ? 'eaVert' : undefined,
  }
}

/** the metroBlob's paragraph indents win over the legacy records when they line up one for one */
function applyMetroIndents(paragraphs: ResolvedParagraph[], blob: Uint8Array): void {
  const indents = metroParagraphIndents(blob)
  // the blob keeps the trailing empty paragraph the legacy text drops
  if (!indents || indents.length < paragraphs.length || indents.length > paragraphs.length + 1)
    return
  for (const [index, paragraph] of paragraphs.entries()) {
    const { marL, indent } = indents[index]!
    if (marL !== undefined) paragraph.leftMargin = marL / EMU_PER_MASTER_UNIT
    if (indent !== undefined) paragraph.indent = paragraph.leftMargin + indent / EMU_PER_MASTER_UNIT
  }
}

/**
 * Turn the '*' placeholder characters the MetaCharacter atoms name into
 * fields: the run holding each position is split so the field is a run of
 * its own (paragraph marks count one character each, like the atoms do).
 */
function markFields(paragraphs: ResolvedParagraph[], fields: TextField[]): void {
  if (fields.length === 0) return
  let offset = 0
  for (const paragraph of paragraphs) {
    const runs: ResolvedRun[] = []
    for (const run of paragraph.runs) {
      const start = offset
      const end = offset + run.text.length
      const hits = fields
        .filter((f) => f.position >= start && f.position < end)
        .sort((a, b) => a.position - b.position)
      if (hits.length === 0) {
        runs.push(run)
      } else {
        let cursor = start
        for (const hit of hits) {
          if (hit.position > cursor)
            runs.push({ ...run, text: run.text.slice(cursor - start, hit.position - start) })
          runs.push({
            ...run,
            text: run.text.slice(hit.position - start, hit.position - start + 1),
            field: hit.type,
          })
          cursor = hit.position + 1
        }
        if (cursor < end) runs.push({ ...run, text: run.text.slice(cursor - start) })
      }
      offset = end
    }
    paragraph.runs = runs
    offset += 1
  }
}

/**
 * The text boxes of a slide (stage 1's flat text list, kept for the
 * `textSequences` view). A real slide nests them under PPDrawing →
 * OfficeArtDgContainer → OfficeArtSpgrContainer → OfficeArtSpContainer →
 * OfficeArtClientTextbox, so the walk follows every record container.
 */
function textBoxesOf(document: Uint8Array, slideRecord: PptRecord): PptSlideText[] {
  const out: PptSlideText[] = []
  collectTextBoxes(document, slideRecord.offset, slideRecord.offset + slideRecord.length, 0, out)
  return out
}

function collectTextBoxes(
  document: Uint8Array,
  from: number,
  to: number,
  depth: number,
  out: PptSlideText[],
): void {
  if (depth > MAX_RECORD_DEPTH) return
  for (const record of walkRecordsLenient(document, from, to)) {
    if (record.type === PPT_TYPES.escherClientTextbox) {
      const text = textOfTextbox(document, record)
      if (text) out.push(text)
      continue
    }
    if (isRecordContainer(record)) {
      collectTextBoxes(document, record.offset, record.offset + record.length, depth + 1, out)
    }
  }
}

/**
 * One ClientTextbox's text. A text box whose text lives in the slide's
 * SlideListWithText carries an OutlineTextRefAtom and no text atom, so it
 * contributes nothing here — its text already came from the outline entries.
 */
function textOfTextbox(document: Uint8Array, record: PptRecord): PptSlideText | null {
  let kind = 4 // TextHeaderEnum "other text"; a box without a header still reads as text
  let text: string | null = null
  for (const inner of walkRecordsLenient(document, record.offset, record.offset + record.length)) {
    if (inner.type === PPT_TYPES.textHeaderAtom) {
      kind = textKindOf(document, inner) ?? kind
      continue
    }
    if (inner.type !== PPT_TYPES.textCharsAtom && inner.type !== PPT_TYPES.textBytesAtom) continue
    text = textOf(document, inner)
  }
  return text === null ? null : { kind, text }
}

function textKindOf(document: Uint8Array, record: PptRecord): number | null {
  const payload = payloadOf(document, record)
  if (payload.length < 4) return null
  return viewOf(payload).getInt32(0, true)
}

function textOf(document: Uint8Array, record: PptRecord): string {
  const payload = payloadOf(document, record)
  return record.type === PPT_TYPES.textCharsAtom
    ? decodeTextChars(payload)
    : decodeTextBytes(payload)
}

function rawTextOf(document: Uint8Array, record: PptRecord): string {
  const payload = payloadOf(document, record)
  return record.type === PPT_TYPES.textCharsAtom ? rawTextChars(payload) : rawTextBytes(payload)
}

function normalizeRaw(raw: string): string {
  return [...raw]
    .map((ch) =>
      ch === '\r' || ch === '\u000b' ? '\n' : ch === '\u001e' || ch === '\u001f' ? ' ' : ch,
    )
    .join('')
}

function viewOf(payload: Uint8Array): DataView {
  return new DataView(payload.buffer, payload.byteOffset, payload.byteLength)
}

/** The record header at a stream offset (or null when it would run off). */
function recordAt(stream: Uint8Array, offset: number): PptRecord | null {
  if (offset < 0 || offset + 8 > stream.length) return null
  const head = stream[offset]! | (stream[offset + 1]! << 8)
  const type = stream[offset + 2]! | (stream[offset + 3]! << 8)
  const length = viewOf(stream.subarray(offset + 4, offset + 8)).getUint32(0, true)
  if (offset + 8 + length > stream.length) return null
  return { version: head & 0x000f, instance: head >> 4, type, length, offset: offset + 8 }
}

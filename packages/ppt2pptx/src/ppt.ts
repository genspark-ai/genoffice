import * as CFB from 'cfb'
import { LegacyPptError } from './errors'
import {
  PPT_TYPES,
  type PptRecord,
  decodeTextBytes,
  decodeTextChars,
  payloadOf,
  walkRecords,
} from './records'

/** Caps mirror the other legacy importers: fail fast, never walk gigabytes. */
const MAX_INPUT_BYTES = 256 * 1024 * 1024
const MAX_STREAM_BYTES = 128 * 1024 * 1024
const MAX_SLIDES = 4096

export interface PptSlideText {
  /** the TextHeaderEnum of the sequence (0 title, 1 body, 2 notes, 4 other…) */
  kind: number
  text: string
}

export interface PptSlide {
  textSequences: PptSlideText[]
}

export interface PptPresentation {
  slides: PptSlide[]
  /** slide surface in master units (576000 per inch) */
  slideWidth: number
  slideHeight: number
}

/**
 * The change directory: persistId → stream offset. PowerPoint's stream holds
 * one generation per save, and the PersistPtr blocks carry the offsets — ids
 * are implicit (the oldest block covers 1..count, the next continues) and a
 * later block overrides an earlier one for the same id. The CurrentUser
 * stream names the live UserEditAtom, whose docPersistId names the live
 * DocumentContainer; a sequential walk would read stale generations.
 */
export interface PersistChain {
  directory: Map<number, number>
  /** the DocumentContainer's persist id, when the chain names one */
  docPersistId: number | null
}

export function buildPersistChain(
  document: Uint8Array,
  currentUser: Uint8Array | null,
): PersistChain {
  const directory = new Map<number, number>()
  let docPersistId: number | null = null
  let editOffset: number | null = null
  if (currentUser) {
    for (const record of walkRecords(currentUser)) {
      if (record.type === PPT_TYPES.currentUserAtom) {
        const payload = payloadOf(currentUser, record)
        if (payload.length >= 12) {
          const v = new DataView(payload.buffer, payload.byteOffset, payload.byteLength)
          // headerToken guards the layout; offsetToCurrentEdit follows the size
          const token = v.getUint32(4, true)
          if (token === 0xe391c05f) editOffset = v.getUint32(8, true)
        }
      }
    }
  }
  // one pass over the presentation stream: every PersistPtr block contributes
  // its offsets under the ids its position implies, and the UserEditAtom the
  // CurrentUser pointed at names the live document
  let nextId = 1
  for (const record of walkRecords(document)) {
    if (record.type === PPT_TYPES.persistPtrIncrementalBlock) {
      const count = record.instance + 1
      if (record.length < count * 4) throw new LegacyPptError('bad_persist')
      const payload = payloadOf(document, record)
      const v = new DataView(payload.buffer, payload.byteOffset, payload.byteLength)
      for (let i = 0; i < count; i++) {
        // offset 0 is a legal stream position (a first slide can open the
        // stream); only past-the-end values are garbage
        const offset = v.getUint32(i * 4, true)
        if (offset >= document.length) continue
        directory.set(nextId + i, offset)
      }
      nextId += count
    }
    if (record.type === PPT_TYPES.userEditAtom && editOffset !== null) {
      if (record.offset - 8 === editOffset) {
        const payload = payloadOf(document, record)
        if (payload.length >= 8) {
          const v = new DataView(payload.buffer, payload.byteOffset, payload.byteLength)
          const docId = v.getUint32(4, true)
          if (docId > 0) docPersistId = docId
        }
      }
    }
  }
  return { directory, docPersistId }
}

/**
 * The presentation: slide size from the DocumentAtom, the slide order from
 * the DocumentContainer's SlideListWithText (persist ids that resolve to a
 * SlideContainer), and each slide's text sequences in place.
 */
export function parsePptPresentation(
  document: Uint8Array,
  currentUser: Uint8Array | null,
): PptPresentation {
  const { directory, docPersistId } = buildPersistChain(document, currentUser)
  // single-generation files (producers without change tracking) walk flat
  const at = (persistId: number): PptRecord | null =>
    directory.size > 0
      ? directory.has(persistId)
        ? recordAt(document, directory.get(persistId)!)
        : null
      : recordAt(document, persistId)

  // the DocumentContainer: a chained file points at it through the last
  // UserEditAtom's docPersistId (the id before the persistPtr started its own
  // numbering); a flat single-generation file is the only DocumentContainer
  let documentRecord: PptRecord | null = null
  if (docPersistId !== null) {
    documentRecord = at(docPersistId)
  }
  if (!documentRecord) {
    for (const record of walkRecords(document)) {
      if (record.type === PPT_TYPES.documentContainer) documentRecord = record
    }
  }
  void at
  if (!documentRecord) throw new LegacyPptError('corrupt', 'no DocumentContainer')

  // DocumentContainer children: the DocumentAtom carries the slide surface,
  // the SlideListWithText(s) carry the slide order
  let slideWidth = 5760000
  let slideHeight = 4320000
  const slideOrder: number[] = []
  const docEnd = documentRecord.offset + documentRecord.length
  for (const record of walkRecords(document, documentRecord.offset, docEnd)) {
    if (record.type === PPT_TYPES.documentAtom && record.length >= 16) {
      // the child's payload offset is absolute in the stream: view it there
      const v = new DataView(document.buffer, document.byteOffset + record.offset, 16)
      slideWidth = v.getUint32(0, true)
      slideHeight = v.getUint32(4, true)
    }
    if (record.type === PPT_TYPES.slideListWithText) {
      const payload = payloadOf(document, record)
      const count = Math.floor(payload.length / 4)
      for (let i = 0; i < count; i++) {
        const v = new DataView(payload.buffer, payload.byteOffset + i * 4, 4)
        const persistId = v.getUint32(0, true)
        if (isSlideContainer(document, at(persistId))) slideOrder.push(persistId)
      }
    }
  }
  // without a SlideListWithText, the persist ids that name SlideContainers in
  // ascending order stand in (their stream order is the deck order then)
  if (slideOrder.length === 0) {
    const ids = [...directory.entries()]
      .filter(([, offset]) => isSlideContainer(document, recordAt(document, offset)))
      .sort((a, b) => a[1] - b[1])
    slideOrder.push(...ids.map(([id]) => id))
  }
  if (slideOrder.length > MAX_SLIDES) {
    throw new LegacyPptError('limits_exceeded', `more than ${MAX_SLIDES} slides`)
  }

  const slides = slideOrder.map((persistId) => {
    const record = at(persistId)
    if (!record || record.type !== PPT_TYPES.slideContainer) {
      throw new LegacyPptError('bad_slide', `persist id ${persistId}`)
    }
    return parseSlide(document, record)
  })
  return { slides, slideWidth, slideHeight }
}

/** The record header at a stream offset (or null when it would run off). */
function recordAt(stream: Uint8Array, offset: number): PptRecord | null {
  if (offset < 0 || offset + 8 > stream.length) return null
  const head = stream[offset]! | (stream[offset + 1]! << 8)
  const type = stream[offset + 2]! | (stream[offset + 3]! << 8)
  const v = new DataView(stream.buffer, stream.byteOffset + offset + 4, 4)
  const length = v.getUint32(0, true)
  if (offset + 8 + length > stream.length) return null
  return { version: head & 0x000f, instance: head >> 4, type, length, offset: offset + 8 }
}

function isSlideContainer(document: Uint8Array, record: PptRecord | null): boolean {
  return record !== null && record.type === PPT_TYPES.slideContainer
}

/** The text sequences of one slide: TextHeaderAtom + TextChars/BytesAtom pairs. */
function parseSlide(document: Uint8Array, slideRecord: PptRecord): PptSlide {
  const slide: PptSlide = { textSequences: [] }
  const end = slideRecord.offset + slideRecord.length
  let pendingKind: number | null = null
  for (const record of walkRecords(document, slideRecord.offset, end)) {
    if (record.type === PPT_TYPES.textHeaderAtom) {
      const payload = payloadOf(document, record)
      pendingKind =
        payload.length >= 4
          ? new DataView(payload.buffer, payload.byteOffset, 4).getInt32(0, true)
          : 4
      continue
    }
    if (
      (record.type === PPT_TYPES.textCharsAtom || record.type === PPT_TYPES.textBytesAtom) &&
      pendingKind !== null
    ) {
      const payload = payloadOf(document, record)
      const text =
        record.type === PPT_TYPES.textCharsAtom
          ? decodeTextChars(payload)
          : decodeTextBytes(payload)
      if (text.trim().length > 0) slide.textSequences.push({ kind: pendingKind, text })
      pendingKind = null
      continue
    }
    // containers between the header and the text atoms recurse (a text run
    // inside an outline text container keeps its header pairing)
    if (record.version === 0xf && record.type !== PPT_TYPES.slideContainer) {
      for (const inner of walkRecords(document, record.offset, record.offset + record.length)) {
        if (inner.type === PPT_TYPES.textHeaderAtom) {
          const payload = payloadOf(document, inner)
          pendingKind =
            payload.length >= 4
              ? new DataView(payload.buffer, payload.byteOffset, 4).getInt32(0, true)
              : 4
          continue
        }
        if (
          (inner.type === PPT_TYPES.textCharsAtom || inner.type === PPT_TYPES.textBytesAtom) &&
          pendingKind !== null
        ) {
          const payload = payloadOf(document, inner)
          const text =
            inner.type === PPT_TYPES.textCharsAtom
              ? decodeTextChars(payload)
              : decodeTextBytes(payload)
          if (text.trim().length > 0) slide.textSequences.push({ kind: pendingKind, text })
          pendingKind = null
        }
      }
    }
  }
  return slide
}

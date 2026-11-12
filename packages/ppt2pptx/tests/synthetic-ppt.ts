import * as CFB from 'cfb'

/**
 * A minimal MS-PPT binary builder: records are laid out the way PowerPoint
 * writes them — a change directory (CurrentUserAtom → UserEditAtom →
 * PersistDirectoryAtom) over one or more generations, slide text nested under
 * PPDrawing/OfficeArt and outline text under SlideListWithText — so the tests
 * exercise the reader against bytes whose content is known exactly.
 */

export interface SlideSpec {
  textSequences: Array<{ kind: number; text: string; atom?: 'chars' | 'bytes' }>
}

export interface BuildOptions {
  /** slide surface in master units (576 per inch); default 5760 × 4320 (10 × 7.5 in) */
  slideSize?: [number, number]
  /** a second generation restating the slides (stale first-gen must be ignored) */
  secondGeneration?: SlideSpec[]
  /**
   * Store each slide's text in the SlideListWithText outline (the real
   * PowerPoint shape for placeholders) instead of in the drawing's
   * OfficeArtClientTextbox.
   */
  outlineTexts?: boolean
  /** write the password-protected CurrentUserAtom headerToken (0xF3D1C4DF) */
  encrypted?: boolean
}

const record = (type: number, payload: Uint8Array, instance = 0, version = 0): Uint8Array => {
  const out = new Uint8Array(8 + payload.length)
  const v = new DataView(out.buffer)
  v.setUint16(0, ((instance & 0xfff) << 4) | (version & 0xf), true)
  v.setUint16(2, type, true)
  v.setUint32(4, payload.length, true)
  out.set(payload, 8)
  return out
}

const u16 = (value: number): Uint8Array => {
  const out = new Uint8Array(2)
  new DataView(out.buffer).setUint16(0, value, true)
  return out
}

const u32 = (value: number): Uint8Array => {
  const out = new Uint8Array(4)
  new DataView(out.buffer).setUint32(0, value >>> 0, true)
  return out
}

const concat = (parts: Uint8Array[]): Uint8Array => {
  const total = parts.reduce((sum, p) => sum + p.length, 0)
  const out = new Uint8Array(total)
  let at = 0
  for (const part of parts) {
    out.set(part, at)
    at += part.length
  }
  return out
}

const textChars = (text: string): Uint8Array =>
  record(0x0fa0, new Uint8Array(Buffer.from(text, 'utf16le')))

const textBytes = (text: string): Uint8Array =>
  record(0x0fa8, new Uint8Array(Buffer.from(text, 'latin1')))

const textHeader = (kind: number): Uint8Array => record(0x0f9f, u32(kind))

/** OfficeArtClientTextbox with a Reference to the slide's outline text entry. */
const outlineTextRef = (index: number): Uint8Array => record(0x0f9e, u32(index))

const slideAtom = (): Uint8Array => record(0x03ef, new Uint8Array(20))

/** SlidePersistAtom: persistIdRef, unused, cTexts, slideId, reserved. */
const slidePersistAtom = (persistIdRef: number, slideId: number, cTexts: number): Uint8Array =>
  record(0x03f3, concat([u32(persistIdRef), u32(0), u32(cTexts), u32(slideId), u32(0)]))

const documentAtom = (width: number, height: number): Uint8Array =>
  record(0x03e9, concat([u32(width), u32(height), u32(width), u32(height)]))

/** one Escher record: the header layout is MS-PPT's (ver/inst, type, length) */
const escher = (type: number, payload: Uint8Array, instance = 0, version = 0): Uint8Array =>
  record(type, payload, instance, version)

const escherContainer = (type: number, children: Uint8Array[]): Uint8Array =>
  escher(type, concat(children), 0, 0xf)

/**
 * A slide's PPDrawing: DgContainer → SpgrContainer → one SpContainer per text
 * sequence → OfficeArtClientTextbox. This is how a real PowerPoint slide nests
 * its text (five levels under the slide container), so the reader's deep walk
 * is what moves these bytes.
 */
function drawing(spec: SlideSpec, outlineTexts: boolean): Uint8Array {
  const shapes = spec.textSequences.map((sequence, index) => {
    const atoms: Uint8Array[] = [textHeader(sequence.kind)]
    if (outlineTexts) atoms.push(outlineTextRef(index))
    else atoms.push(sequence.atom === 'bytes' ? textBytes(sequence.text) : textChars(sequence.text))
    const textbox = escher(0xf00d, concat(atoms))
    return escherContainer(0xf004, [textbox])
  })
  const spgr = escherContainer(0xf003, shapes)
  const dg = escherContainer(0xf002, [spgr])
  return escherContainer(0x040c, [dg])
}

function slideContainer(spec: SlideSpec, outlineTexts: boolean): Uint8Array {
  return record(0x03ee, concat([slideAtom(), drawing(spec, outlineTexts)]), 0, 0xf)
}

/** The outline text pairs SlideListWithText carries for one slide. */
function outlineText(spec: SlideSpec): Uint8Array[] {
  const parts: Uint8Array[] = []
  for (const sequence of spec.textSequences) {
    parts.push(textHeader(sequence.kind))
    parts.push(sequence.atom === 'bytes' ? textBytes(sequence.text) : textChars(sequence.text))
  }
  return parts
}

/** The PowerPoint Document stream + Current User stream, chained the way PowerPoint saves. */
export function buildPptStreams(
  slideSpecs: SlideSpec[],
  options: BuildOptions = {},
): { document: Uint8Array; currentUser: Uint8Array } {
  const slideSize = options.slideSize ?? [5760, 4320]
  const outlineTexts = options.outlineTexts === true
  const generations: SlideSpec[][] = options.secondGeneration
    ? [slideSpecs, options.secondGeneration]
    : [slideSpecs]

  // lay every part out in stream order, recording each persist id's offset
  const persistOffsets = new Map<number, number>()
  const parts: Uint8Array[] = []
  let offset = 0
  let nextId = 1
  let nextSlideId = 256
  let lastUserEditOffset = 0

  const push = (bytes: Uint8Array): void => {
    parts.push(bytes)
    offset += bytes.length
  }

  for (const [generation, specs] of generations.entries()) {
    const slides: { persistId: number; slideId: number }[] = []
    for (const spec of specs) {
      const start = offset
      const persistId = nextId++
      const slideId = nextSlideId++
      push(slideContainer(spec, outlineTexts))
      slides.push({ persistId, slideId })
      persistOffsets.set(persistId, start)
    }

    // DocumentContainer: DocumentAtom + SlideListWithText (instance 0 = slides)
    const listParts: Uint8Array[] = []
    for (const [index, spec] of specs.entries()) {
      const slide = slides[index]!
      listParts.push(slidePersistAtom(slide.persistId, slide.slideId, spec.textSequences.length))
      if (outlineTexts) listParts.push(...outlineText(spec))
    }
    const docStart = offset
    push(
      record(
        0x03e8,
        concat([
          documentAtom(slideSize[0], slideSize[1]),
          record(0x0ff0, concat(listParts), 0, 0xf),
          record(0x03ea, new Uint8Array(0)),
        ]),
        0,
        0xf,
      ),
    )
    const docPersistId = nextId++
    persistOffsets.set(docPersistId, docStart)

    // PersistDirectoryAtom: one PersistDirectoryEntry per persist id (cPersist=1)
    const persistParts: Uint8Array[] = []
    for (const id of [...slides.map((slide) => slide.persistId), docPersistId]) {
      persistParts.push(u32(id | (1 << 20)), u32(persistOffsets.get(id)!))
    }
    const persistStart = offset
    push(record(0x1772, concat(persistParts)))

    // UserEditAtom (28-byte payload: MS-PPT 2.3.3 field order)
    const editStart = offset
    push(
      record(
        0x0ff5,
        concat([
          u32(0), // lastSlideIdRef
          u16(0x00f4), // version
          new Uint8Array([0, 3]), // minorVersion, majorVersion
          u32(generation === 0 ? 0 : lastUserEditOffset), // offsetLastEdit
          u32(persistStart), // offsetPersistDirectory
          u32(docPersistId), // docPersistIdRef
          u32(nextId), // persistIdSeed
          u16(0),
          u16(0), // lastView, unused
        ]),
      ),
    )
    lastUserEditOffset = editStart
  }

  const currentUser = record(
    0x0ff6,
    concat([
      u32(0x14), // size
      u32(options.encrypted ? 0xf3d1c4df : 0xe391c05f), // headerToken
      u32(lastUserEditOffset), // offsetToCurrentEdit
      u16(0), // lenUserName
      u16(0x03f4), // docFileVersion
      new Uint8Array([3, 0, 0]), // majorVersion, minorVersion, unused
      u16(0), // reserved
    ]),
  )

  return { document: concat(parts), currentUser }
}

export function buildPptContainer(document: Uint8Array, currentUser: Uint8Array): Uint8Array {
  const cfb = CFB.utils.cfb_new({ root: 'Root Entry' })
  CFB.utils.cfb_add(cfb, '/PowerPoint Document', Buffer.from(document))
  CFB.utils.cfb_add(cfb, '/Current User', Buffer.from(currentUser))
  const out = CFB.write(cfb, { type: 'buffer' }) as Buffer
  return new Uint8Array(out.buffer, out.byteOffset, out.byteLength)
}

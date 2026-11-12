import * as CFB from 'cfb'

/**
 * A minimal MS-PPT binary builder: records are laid out the way PowerPoint
 * writes them — a change directory (CurrentUserAtom → UserEditAtom →
 * PersistDirectoryAtom) over one or more generations, slide text nested under
 * PPDrawing/OfficeArt and outline text under SlideListWithText — so the tests
 * exercise the reader against bytes whose content is known exactly.
 */

export interface TextSpec {
  kind: number
  text: string
  atom?: 'chars' | 'bytes'
  /** OfficeArtClientAnchor in master units: top, left, right, bottom */
  anchor?: [number, number, number, number]
  /** OfficeArt properties (pid → value) on the shape */
  props?: Record<number, number>
  /** OfficeArt complex properties (pid → payload) on the shape */
  complex?: Record<number, Uint8Array>
  /** msosp shape type (FSP instance); default 202 text box */
  shapeType?: number
  /** a StyleTextPropAtom: one paragraph run and one character run over the whole text */
  style?: {
    align?: number
    bold?: boolean
    shadow?: boolean
    size?: number
    colorRgb?: number
    bullet?: boolean
    /** a bullet colour (0xBBGGRR) and whether fBulletHasColor says the bullet uses it */
    bulletColorRgb?: number
    bulletHasColor?: boolean
  }
  /** MetaCharacter atoms (record type + the position of the '*' they stand for) */
  metaChars?: Array<{ type: number; position: number }>
  /** a TextRulerAtom payload to follow the style runs */
  ruler?: Uint8Array
  /** a metroBlob (OPT 0x03a9) package on a tertiary OPT */
  metroBlob?: Uint8Array
  /** a PP9 StyleTextProp9Atom numbering the (single) paragraph run */
  autoNumber?: { scheme?: number; startAt?: number }
}

export interface PictureSpec {
  anchor: [number, number, number, number]
  /** 1-based index into the Pictures stream */
  pib: number
  /** further OfficeArt properties (pid → value) on the picture */
  props?: Record<number, number>
}

export interface GroupSpec {
  /** the group's slide-space anchor (top, left, right, bottom) */
  anchor: [number, number, number, number]
  /** the group coordinate space (left, top, right, bottom) its members are anchored in */
  bounds: [number, number, number, number]
  /** degrees clockwise */
  rotation?: number
  flipH?: boolean
  members: Array<{ childAnchor: [number, number, number, number]; text: string; rotation?: number }>
}

export interface SlideSpec {
  textSequences: TextSpec[]
  pictures?: PictureSpec[]
  groups?: GroupSpec[]
  /** a background shape with this solid fill colour (0xBBGGRR) */
  backgroundRgb?: number
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
  /** PNG blips for the Pictures stream, in pib order */
  pictures?: Uint8Array[]
  /** with outlineTexts: write each sequence's MetaCharacter atoms into the outline, and no shapes at all */
  outlineMetaChars?: boolean
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
const i16 = (value: number): Uint8Array => {
  const out = new Uint8Array(2)
  new DataView(out.buffer).setInt16(0, value, true)
  return out
}

/** OfficeArtFSP: spid + flags, the shape type in the instance */
const fsp = (shapeType: number, flags: number, spid = 1024): Uint8Array =>
  escher(0xf00a, concat([u32(spid), u32(flags)]), shapeType, 2)

/** OfficeArtFOPT: the simple properties, then the complex ones (fComplex set, value = payload length) */
const fopt = (
  props: Record<number, number>,
  complex: Record<number, Uint8Array> = {},
): Uint8Array => {
  const entries = Object.entries(props).map(([pid, value]) =>
    concat([u16(Number(pid)), u32(value)]),
  )
  const payloads: Uint8Array[] = []
  for (const [pid, payload] of Object.entries(complex)) {
    entries.push(concat([u16(Number(pid) | 0x8000), u32(payload.length)]))
    payloads.push(payload)
  }
  return escher(0xf00b, concat([...entries, ...payloads]), entries.length, 3)
}

/** OfficeArtTertiaryFOPT holding one complex property */
const complexFopt = (pid: number, bytes: Uint8Array): Uint8Array =>
  escher(0xf122, concat([u16(pid | 0x8000), u32(bytes.length), bytes]), 1, 3)

const clientAnchor = (anchor: [number, number, number, number]): Uint8Array =>
  escher(0xf010, concat(anchor.map(i16)))

/** StyleTextPropAtom: one paragraph run + one character run covering text + mark */
function styleTextProp(length: number, style: NonNullable<TextSpec['style']>): Uint8Array {
  const paragraph: Uint8Array[] = [u32(length + 1), u16(0)]
  let pMask = 0
  const pFields: Uint8Array[] = []
  if (style.bullet !== undefined || style.bulletHasColor !== undefined) {
    pMask |=
      (style.bullet !== undefined ? 0x0001 : 0) | (style.bulletHasColor !== undefined ? 0x0004 : 0)
    pFields.push(u16((style.bullet ? 1 : 0) | (style.bulletHasColor ? 4 : 0)))
  }
  if (style.bulletColorRgb !== undefined) {
    pMask |= 0x0020
    pFields.push(u32(style.bulletColorRgb | 0xfe000000))
  }
  if (style.align !== undefined) {
    pMask |= 0x0800
    pFields.push(u16(style.align))
  }
  paragraph.push(u32(pMask), ...pFields)
  const chars: Uint8Array[] = [u32(length + 1)]
  let cMask = 0
  const cFields: Uint8Array[] = []
  if (style.bold !== undefined || style.shadow !== undefined) {
    cMask |= (style.bold !== undefined ? 0x0001 : 0) | (style.shadow !== undefined ? 0x0010 : 0)
    cFields.push(u16((style.bold ? 1 : 0) | (style.shadow ? 0x10 : 0)))
  }
  if (style.size !== undefined) {
    cMask |= 0x20000
    cFields.push(u16(style.size))
  }
  if (style.colorRgb !== undefined) {
    cMask |= 0x40000
    cFields.push(u32(style.colorRgb | 0xfe000000))
  }
  chars.push(u32(cMask), ...cFields)
  return record(0x0fa1, concat([...paragraph, ...chars]))
}

const cString = (text: string): Uint8Array =>
  record(0x0fba, new Uint8Array(Buffer.from(text, 'utf16le')))

/** one StyleTextProp9 entry: TextPFException9 with the scheme, empty CF9 and SI */
const styleTextProp9 = (autoNumber: NonNullable<TextSpec['autoNumber']>): Uint8Array =>
  record(
    0x0fac,
    concat([
      u32(0x03000000),
      u16(1),
      u16(autoNumber.scheme ?? 3),
      u16(autoNumber.startAt ?? 1),
      u32(0),
      u32(0),
    ]),
  )

/** ProgTags -> ProgBinaryTag "___PPT9" -> PP9 extension around these records */
const pp9Tag = (children: Uint8Array[]): Uint8Array =>
  record(
    0x1388,
    record(0x138a, concat([cString('___PPT9'), record(0x138b, concat(children), 0, 0xf)]), 0, 0xf),
    0,
    0xf,
  )

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
    else {
      atoms.push(sequence.atom === 'bytes' ? textBytes(sequence.text) : textChars(sequence.text))
      if (sequence.style) atoms.push(styleTextProp(sequence.text.length, sequence.style))
      for (const meta of sequence.metaChars ?? []) atoms.push(record(meta.type, u32(meta.position)))
      if (sequence.ruler) atoms.push(record(0x0fa6, sequence.ruler))
    }
    const textbox = escher(0xf00d, concat(atoms))
    const parts: Uint8Array[] = []
    if (sequence.anchor || sequence.props || sequence.shapeType !== undefined) {
      parts.push(fsp(sequence.shapeType ?? 202, 0xa00, 1024 + index))
      parts.push(fopt(sequence.props ?? {}, sequence.complex))
      if (sequence.metroBlob) parts.push(complexFopt(0x03a9, sequence.metroBlob))
      if (sequence.anchor) parts.push(clientAnchor(sequence.anchor))
    }
    if (sequence.autoNumber && !outlineTexts)
      parts.push(escher(0xf011, pp9Tag([styleTextProp9(sequence.autoNumber)])))
    parts.push(textbox)
    return escherContainer(0xf004, parts)
  })
  for (const [index, picture] of (spec.pictures ?? []).entries()) {
    shapes.push(
      escherContainer(0xf004, [
        fsp(75, 0xa00, 2048 + index),
        fopt({ 0x0104: picture.pib, 0x01bf: 0x00100000, 0x01ff: 0x00080000, ...picture.props }),
        clientAnchor(picture.anchor),
      ]),
    )
  }
  for (const [index, group] of (spec.groups ?? []).entries()) {
    const rect32 = (r: [number, number, number, number]): Uint8Array => concat(r.map((v) => u32(v)))
    const rotationProp = (deg?: number): Record<number, number> =>
      deg === undefined ? {} : { 0x0004: Math.round(deg * 65536) >>> 0 }
    const self = escherContainer(0xf004, [
      escher(0xf009, rect32(group.bounds), 0, 1),
      fsp(0, 0x201 | (group.flipH ? 0x40 : 0), 3000 + index * 100),
      fopt(rotationProp(group.rotation)),
      clientAnchor(group.anchor),
    ])
    const members = group.members.map((member, m) =>
      escherContainer(0xf004, [
        fsp(202, 0xa02, 3001 + index * 100 + m),
        fopt(rotationProp(member.rotation)),
        escher(0xf00f, rect32(member.childAnchor)),
        escher(0xf00d, concat([textHeader(4), textChars(member.text)])),
      ]),
    )
    shapes.push(escherContainer(0xf003, [self, ...members]))
  }
  const spgr = escherContainer(0xf003, shapes)
  const children: Uint8Array[] = [spgr]
  if (spec.backgroundRgb !== undefined) {
    children.push(
      escherContainer(0xf004, [
        fsp(1, 0xc00, 4095),
        fopt({ 0x0181: spec.backgroundRgb, 0x01bf: 0x00100010 }),
      ]),
    )
  }
  const dg = escherContainer(0xf002, children)
  return escherContainer(0x040c, [dg])
}

function slideContainer(spec: SlideSpec, outlineTexts: boolean, noShapes = false): Uint8Array {
  const body = noShapes
    ? drawing({ ...spec, textSequences: [] }, outlineTexts)
    : drawing(spec, outlineTexts)
  return record(0x03ee, concat([slideAtom(), body]), 0, 0xf)
}

/** The outline text pairs SlideListWithText carries for one slide. */
function outlineText(spec: SlideSpec, metaChars = false): Uint8Array[] {
  const parts: Uint8Array[] = []
  for (const sequence of spec.textSequences) {
    parts.push(textHeader(sequence.kind))
    if (sequence.text.length === 0) continue
    parts.push(sequence.atom === 'bytes' ? textBytes(sequence.text) : textChars(sequence.text))
    if (sequence.style) parts.push(styleTextProp(sequence.text.length, sequence.style))
    if (metaChars)
      for (const meta of sequence.metaChars ?? []) parts.push(record(meta.type, u32(meta.position)))
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
      push(slideContainer(spec, outlineTexts, options.outlineMetaChars === true))
      slides.push({ persistId, slideId })
      persistOffsets.set(persistId, start)
    }

    // DocumentContainer: DocumentAtom + SlideListWithText (instance 0 = slides)
    const listParts: Uint8Array[] = []
    for (const [index, spec] of specs.entries()) {
      const slide = slides[index]!
      listParts.push(slidePersistAtom(slide.persistId, slide.slideId, spec.textSequences.length))
      if (outlineTexts) listParts.push(...outlineText(spec, options.outlineMetaChars === true))
    }
    // the outline texts' PP9 numbering: OutlineTextProps9 entries under the List container
    const props9: Uint8Array[] = []
    for (const [index, spec] of specs.entries()) {
      for (const sequence of spec.textSequences) {
        if (!outlineTexts || !sequence.autoNumber) continue
        props9.push(
          record(0x0faf, concat([u32(slides[index]!.slideId), u32(sequence.kind)]), sequence.kind),
          styleTextProp9(sequence.autoNumber),
        )
      }
    }
    const list = props9.length
      ? [record(0x07d0, pp9Tag([record(0x0fae, concat(props9), 0, 0xf)]), 0, 0xf)]
      : []
    const docStart = offset
    push(
      record(
        0x03e8,
        concat([
          documentAtom(slideSize[0], slideSize[1]),
          record(0x0ff0, concat(listParts), 0, 0xf),
          ...list,
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

/** the Pictures stream: one OfficeArtBlipPNG (or BlipDIB for a `{ dib }` entry) per picture: 16-byte UID, tag byte, bytes */
export function buildPicturesStream(pictures: Array<Uint8Array | { dib: Uint8Array }>): Uint8Array {
  return concat(
    pictures.map((picture) =>
      picture instanceof Uint8Array
        ? escher(0xf01e, concat([new Uint8Array(16), new Uint8Array([0xff]), picture]), 0x6e0)
        : escher(0xf01f, concat([new Uint8Array(16), new Uint8Array([0xff]), picture.dib]), 0x7a8),
    ),
  )
}

/** an 8×8 1-bpp DIB (BITMAPINFOHEADER, white then black palette entry) from top-down rows of black pixels */
export function buildPatternDib(blackRows: number[]): Uint8Array {
  const header = concat([
    u32(40),
    u32(8),
    u32(8),
    u16(1),
    u16(1),
    u32(0),
    u32(32),
    u32(2835),
    u32(2835),
    u32(2),
    u32(2),
  ])
  const palette = new Uint8Array([0xff, 0xff, 0xff, 0, 0, 0, 0, 0])
  const pixels = new Uint8Array(32)
  blackRows.forEach((row, y) => {
    pixels[(7 - y) * 4] = row
  })
  return concat([header, palette, pixels])
}

export function buildPptContainer(
  document: Uint8Array,
  currentUser: Uint8Array,
  pictures?: Uint8Array,
): Uint8Array {
  const cfb = CFB.utils.cfb_new({ root: 'Root Entry' })
  CFB.utils.cfb_add(cfb, '/PowerPoint Document', Buffer.from(document))
  CFB.utils.cfb_add(cfb, '/Current User', Buffer.from(currentUser))
  if (pictures) CFB.utils.cfb_add(cfb, '/Pictures', Buffer.from(pictures))
  const out = CFB.write(cfb, { type: 'buffer' }) as Buffer
  return new Uint8Array(out.buffer, out.byteOffset, out.byteLength)
}

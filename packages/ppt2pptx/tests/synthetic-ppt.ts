import * as CFB from 'cfb'

/**
 * A minimal MS-PPT binary builder: records are laid out the way PowerPoint
 * writes them — a change directory (CurrentUserAtom → UserEditAtom →
 * PersistDirectoryAtom) over one or more generations, slide text nested under
 * PPDrawing/OfficeArt and outline text under SlideListWithText — so the tests
 * exercise the reader against bytes whose content is known exactly.
 */

export interface CharRunSpec {
  count: number
  bold?: boolean
  italic?: boolean
  underline?: boolean
  size?: number
  color?: [number, number, number]
}

export interface ShapeSpec {
  /** [left, top, right, bottom] in master units; absent = no anchor */
  rect?: [number, number, number, number]
  /** msospt* shape type (1 rect, 3 ellipse, 75 picture, 202 text box) */
  shapeType?: number
  text?: string
  kind?: number
  atom?: 'chars' | 'bytes'
  /** text lives in SlideListWithText; the shape carries an OutlineTextRefAtom */
  outlineRef?: boolean
  blipIndex?: number
  fillColor?: number
  lineColor?: number
  flags?: number
  charRuns?: CharRunSpec[]
  /** TextAlignmentEnum: 0 left, 1 center, 2 right, 3 justify */
  paragraphAlign?: number
  /** clockwise degrees, written as OPT pid 0x0004 (16.16 fixed point) */
  rotation?: number
  /** OEPlaceholderAtom.placeholderId; 1 = master title, 13 = title, 14 = body */
  placeholderId?: number
  /** emit as a real Escher group with mapped child coordinates */
  group?: GroupSpec
}

export interface GroupSpec {
  /** FSPGR child coordinate space [left, top, right, bottom] */
  bounds: [number, number, number, number]
  /** the group's own anchor on the slide */
  rect: [number, number, number, number]
  children: ShapeSpec[]
}

export interface MasterSpec {
  shapes: ShapeSpec[]
}

export interface SlideSpec {
  textSequences: Array<{ kind: number; text: string; atom?: 'chars' | 'bytes' }>
  /** stage 3: explicit Escher shapes; absent = stage-1 text boxes */
  shapes?: ShapeSpec[]
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
  /** PNG bytes, one per 1-based blip index, embedded in a global BStore */
  images?: Uint8Array[]
  /** a MainMasterContainer whose anchored shapes the deck inherits */
  master?: MasterSpec
  /** FontCollection typeface names, in font-ref order */
  fonts?: string[]
  /** TextMasterStyleAtom defaults, one level per entry */
  textMasterStyles?: Array<{ textType: number; sizePt: number; fontRef?: number }>
  /** embedded metafile blips (uncompressed), one per 1-based blip index */
  metafiles?: Array<{ kind: 'emf' | 'wmf'; data: Uint8Array }>
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

const i32 = (value: number): Uint8Array => {
  const out = new Uint8Array(4)
  new DataView(out.buffer).setInt32(0, value, true)
  return out
}

const u8 = (...values: number[]): Uint8Array => new Uint8Array(values)

/** one OfficeArtFOPTE: pid (u16, complex bit clear) + value (u32) */
const optEntry = (pid: number, value: number): Uint8Array => {
  const out = new Uint8Array(6)
  const v = new DataView(out.buffer)
  v.setUint16(0, pid, true)
  v.setUint32(2, value, true)
  return out
}

/** StyleTextPropAtom: one paragraph run (optional alignment) + character runs. */
function styleTextProp(spec: ShapeSpec, charCount: number): Uint8Array {
  const paragraph: Uint8Array[] = [u32(charCount), u16(0)]
  if (spec.paragraphAlign === undefined) {
    paragraph.push(u32(0)) // PFMasks: no optional fields
  } else {
    paragraph.push(u32(1 << 11), u16(spec.paragraphAlign)) // PFMasks.align + textAlignment
  }
  const characters: Uint8Array[] = []
  for (const run of spec.charRuns ?? []) {
    const style = (run.bold ? 1 : 0) | (run.italic ? 2 : 0) | (run.underline ? 4 : 0)
    let mask = 0
    if (style !== 0) mask |= 0x07 // bold / italic / underline -> fontStyle exists
    if (run.size !== undefined) mask |= 1 << 17
    if (run.color !== undefined) mask |= 1 << 18
    const fields: Uint8Array[] = [u32(run.count), u32(mask)]
    if (style !== 0) fields.push(u16(style))
    if (run.size !== undefined) fields.push(u16(run.size))
    if (run.color !== undefined) fields.push(u8(...run.color, 0)) // ColorIndexStruct
    characters.push(...fields)
  }
  return record(0x0fa1, concat([...paragraph, ...characters]), charCount)
}

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
function shapeContainer(spec: ShapeSpec, index: number): Uint8Array {
  const children: Uint8Array[] = []
  const shapeType = spec.shapeType ?? 202
  const flags = spec.flags ?? 0
  children.push(escher(0xf00a, concat([u32(index + 2), u32(flags)]), shapeType))
  if (spec.rect) children.push(escher(0xf00f, concat(spec.rect.map(i32))))
  const opt: Uint8Array[] = []
  if (spec.rotation !== undefined)
    opt.push(optEntry(0x0004, Math.round(spec.rotation * 65536) >>> 0))
  if (spec.blipIndex !== undefined) opt.push(optEntry(0x0104, spec.blipIndex))
  if (spec.fillColor !== undefined) opt.push(optEntry(0x0181, spec.fillColor))
  if (spec.lineColor !== undefined) opt.push(optEntry(0x01c0, spec.lineColor))
  if (opt.length > 0) children.push(escher(0xf00b, concat(opt), opt.length))
  if (spec.text !== undefined) {
    const atoms: Uint8Array[] = [textHeader(spec.kind ?? 4)]
    if (spec.outlineRef) atoms.push(outlineTextRef(index))
    else {
      atoms.push(spec.atom === 'bytes' ? textBytes(spec.text) : textChars(spec.text))
      if (spec.charRuns || spec.paragraphAlign !== undefined) {
        atoms.push(styleTextProp(spec, [...spec.text].length))
      }
    }
    children.push(escher(0xf00d, concat(atoms)))
  }
  if (spec.placeholderId !== undefined) {
    const placeholder = record(0x0bc3, concat([u32(0), u8(spec.placeholderId, 0), u16(0)]))
    children.push(escherContainer(0xf011, [placeholder]))
  }
  return escherContainer(0xf004, children)
}

/** A real Escher group: first SpContainer is the group (FSPGR + anchor), the
 *  rest are members whose anchors live in the group's coordinate space. */
function groupContainer(group: GroupSpec, index: number): Uint8Array {
  const self = escherContainer(0xf004, [
    escher(0xf00a, concat([u32(index + 1000), u32(0x5)]), 0), // group | patriarch
    escher(0xf009, concat(group.bounds.map(i32))),
    escher(0xf00f, concat(group.rect.map(i32))),
  ])
  // flat, like a real PowerPoint spgrContainer: the group shape is its first
  // SpContainer and the members follow it
  return concat([self, ...group.children.map(shapeContainer)])
}

/** A slide's PPDrawing: DgContainer -> SpgrContainer -> SpContainers. */
function drawing(spec: SlideSpec, outlineTexts: boolean): Uint8Array {
  const shapes: Uint8Array[] = []
  if (spec.shapes) {
    for (const [index, shape] of spec.shapes.entries()) {
      shapes.push(shape.group ? groupContainer(shape.group, index) : shapeContainer(shape, index))
    }
  } else {
    for (const [index, sequence] of spec.textSequences.entries()) {
      const atoms: Uint8Array[] = [textHeader(sequence.kind)]
      if (outlineTexts) atoms.push(outlineTextRef(index))
      else
        atoms.push(sequence.atom === 'bytes' ? textBytes(sequence.text) : textChars(sequence.text))
      shapes.push(escherContainer(0xf004, [escher(0xf00d, concat(atoms))]))
    }
  }
  const spgr = escherContainer(0xf003, shapes)
  const dg = escherContainer(0xf002, [spgr])
  return escherContainer(0x040c, [dg])
}

/** One embedded OfficeArtBlipPNG: 16-byte UID, a tag byte, then the PNG data. */
function embeddedPng(png: Uint8Array): Uint8Array {
  return escher(0xf01e, concat([new Uint8Array(16), u8(0), png]))
}

/** An uncompressed OfficeArtBlipEMF/WMF: UID, 34-byte metafile header, data. */
function embeddedMetafile(kind: 'emf' | 'wmf', data: Uint8Array): Uint8Array {
  const header = concat([
    new Uint8Array(28),
    u32(data.length), // cbSave
    u8(1, 0), // compression = uncompressed, filter
  ])
  return escher(kind === 'emf' ? 0xf01a : 0xf01b, concat([new Uint8Array(16), header, data]))
}

/** OfficeArtDggContainer with a BStore of FBSE entries carrying embedded blips. */
function drawingGroup(
  images: Uint8Array[],
  metafiles: Array<{ kind: 'emf' | 'wmf'; data: Uint8Array }>,
): Uint8Array {
  const blips = [
    ...images.map(embeddedPng),
    ...metafiles.map((m) => embeddedMetafile(m.kind, m.data)),
  ]
  const entries = blips.map((blip) => {
    const header = concat([
      u8(0, 0), // btWin32, btMacOS
      new Uint8Array(16), // rgbUid
      u16(0), // tag
      u32(blip.length), // size
      u32(1), // cRef
      u32(0), // foDelay (no Pictures stream)
      u8(0, 0, 0, 0), // unused1, cbName, unused2, unused3
    ])
    return escher(0xf007, concat([header, blip]))
  })
  return escherContainer(0xf000, [escherContainer(0xf001, entries)])
}

function slideContainer(spec: SlideSpec, outlineTexts: boolean): Uint8Array {
  return record(0x03ee, concat([slideAtom(), drawing(spec, outlineTexts)]), 0, 0xf)
}

/** FontCollection: FontEntityAtom (0x0FB7) with a 64-byte UTF-16 lfFaceName. */
function fontCollection(fonts: string[]): Uint8Array {
  return record(
    0x07d5,
    concat(
      fonts.map((name) => {
        const face = new Uint8Array(64)
        const units = Buffer.from(name.slice(0, 31), 'utf16le')
        face.set(units.subarray(0, 62))
        return record(0x0fb7, concat([face, u8(0, 0, 0, 0)]))
      }),
    ),
    0,
    0xf,
  )
}

/** TextMasterStyleAtom: one level whose CF exception sets size and font ref. */
function textMasterStyle(spec: { textType: number; sizePt: number; fontRef?: number }): Uint8Array {
  const withFont = spec.fontRef !== undefined
  const mask = (1 << 17) | (withFont ? 1 << 22 : 0)
  const fields: Uint8Array[] = [u32(0), u32(mask)] // PFMasks=0, CFMasks
  if (withFont) fields.push(u16(spec.fontRef!)) // ansiFontRef precedes fontSize
  fields.push(u16(spec.sizePt))
  return record(0x0fa3, concat([u16(1), ...fields]), spec.textType)
}

/** MainMasterContainer: the master's own PPDrawing, inherited by every slide. */
function masterContainer(master: MasterSpec): Uint8Array {
  const shapes = master.shapes.map(shapeContainer)
  const spgr = escherContainer(0xf003, shapes)
  const dg = escherContainer(0xf002, [spgr])
  return record(0x03f8, concat([slideAtom(), escherContainer(0x040c, [dg])]), 0, 0xf)
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
  const images = options.images ?? []
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
          ...(options.fonts ? [fontCollection(options.fonts)] : []),
          ...(options.textMasterStyles ?? []).map(textMasterStyle),
          ...(options.master ? [masterContainer(options.master)] : []),
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

  // the global drawing group is a top-level record; the reader scans for its
  // BStore independently of the change directory
  const metafiles = options.metafiles ?? []
  if (images.length > 0 || metafiles.length > 0) push(drawingGroup(images, metafiles))

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

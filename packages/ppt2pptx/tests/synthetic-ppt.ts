import * as CFB from 'cfb'

/**
 * A minimal MS-PPT binary builder: records written by hand, so the tests
 * exercise the reader against bytes whose content is known exactly.
 */

export interface SlideSpec {
  textSequences: Array<{ kind: number; text: string }>
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

const textHeader = (kind: number): Uint8Array => record(0x0f9f, u32(kind))

const documentAtom = (width = 5760000, height = 4320000): Uint8Array =>
  record(0x03e9, concat([u32(width), u32(height), u32(0), u32(0)]))

/** one slide's text records as a flat run (the container wrapper the builder adds) */
function slideContainer(spec: SlideSpec): Uint8Array {
  const parts: Uint8Array[] = [record(0x03ef, concat([u32(0), u32(0), u32(0), u32(0), u32(0)]))]
  for (const sequence of spec.textSequences) {
    parts.push(textHeader(sequence.kind), textChars(sequence.text))
  }
  return record(0x03ee, concat(parts), 0, 0xf)
}

export interface BuildOptions {
  slideSize?: [number, number]
  /** a second generation restating the slides (stale first-gen must be ignored) */
  secondGeneration?: SlideSpec[]
}

/** The PowerPoint Document stream + Current User stream, chained the way
    PowerPoint saves: CurrentUserAtom → offsetToCurrentEdit → UserEditAtom →
    PersistPtr block → the live generation's containers. Persist ids are
    implicit: the oldest block covers 1..count, the next block continues. */
export function buildPptStreams(
  slideSpecs: SlideSpec[],
  options: BuildOptions = {},
): { document: Uint8Array; currentUser: Uint8Array } {
  const slideSize = options.slideSize ?? [5760000, 4320000]
  const generations: SlideSpec[][] = options.secondGeneration
    ? [slideSpecs, options.secondGeneration]
    : [slideSpecs]

  // lay every part out in stream order, recording the persist mapping
  const parts: Uint8Array[] = []
  const persistDirectory = new Map<number, number>()
  let offset = 0
  let nextId = 1
  let lastUserEditOffset = 0

  const push = (record: Uint8Array): void => {
    parts.push(record)
    offset += record.length
  }

  generations.forEach((specs, genIndex) => {
    const slideIds: number[] = []
    for (const spec of specs) {
      const start = offset
      push(slideContainer(spec))
      persistDirectory.set(nextId, start)
      slideIds.push(nextId)
      nextId += 1
    }
    const docStart = offset
    const docRecord = record(
      0x03e8,
      concat([
        documentAtom(slideSize[0], slideSize[1]),
        record(0x0ff0, concat(slideIds.map((id) => u32(id))), 0, 0xf),
      ]),
      0,
      0xf,
    )
    push(docRecord)
    const docPersistId = nextId
    persistDirectory.set(docPersistId, docStart)
    nextId += 1

    // the PersistPtr block covers this generation's ids; its payload is the
    // plain offset array (the ids are implicit: the previous blocks' counts)
    const entries = [...slideIds, docPersistId]
    const startId = nextId - entries.length
    void startId
    const ptr = record(
      0x1772,
      concat(entries.map((id) => u32(persistDirectory.get(id)!))),
      entries.length - 1,
    )
    push(ptr)

    const editStart = offset
    const userEdit = record(
      0x0ff5,
      concat([
        u32(docPersistId), // lastEditId
        u32(docPersistId), // docPersistId
        u32(nextId), // persistIdSeed
        u16(1), // lastView
        u16(0), // unused
        u32(genIndex === 0 ? 0 : lastUserEditOffset), // offsetToPrevEdit
      ]),
    )
    push(userEdit)
    lastUserEditOffset = editStart
  })

  const currentUser = record(
    0x0ff6,
    concat([
      u32(0x14), // size
      u32(0xe391c05f), // headerToken
      u32(lastUserEditOffset), // offsetToCurrentEdit
      u32(0x30), // lenUserEdit
      u32(0x0000000a),
      u16(0x000a),
      u16(0),
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

/**
 * Hostile input and hostile drawing for the legacy .ppt converter. The
 * contract under attack: every malformed file either converts or fails with
 * a LegacyPptError code, nothing hangs, and no non-finite number ever
 * reaches the .pptx the writer emits.
 *
 * The byte cases corrupt a healthy deck built by synthetic-ppt (fixed
 * surgery plus seeded flips); the drawing cases craft Escher records the
 * builder cannot express — nested groups, degenerate and exploding group
 * spaces, extreme anchors — through SlideSpec.rawDrawing.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  convertPptToPptx,
  LEGACY_PPT_ERROR_CODES,
  LegacyPptError,
  type LegacyPptErrorCode,
} from '../src/index'
import { buildPptContainer, buildPptStreams, type SlideSpec } from './synthetic-ppt'

/* ---------- the contract ---------- */

/** every conversion lands here: a real result, or a coded LegacyPptError */
async function outcome(bytes: Uint8Array): Promise<'ok' | LegacyPptErrorCode> {
  try {
    const result = await convertPptToPptx(bytes)
    expect(result.slideCount).toBeGreaterThanOrEqual(0)
    expect(result.bytes.byteLength).toBeGreaterThan(0)
    return 'ok'
  } catch (error) {
    expect(error).toBeInstanceOf(LegacyPptError)
    const code = (error as LegacyPptError).code
    expect(LEGACY_PPT_ERROR_CODES).toContain(code)
    return code
  }
}

/** like outcome, and on success every slide's geometry must be finite */
async function cleanOutcome(bytes: Uint8Array): Promise<'ok' | LegacyPptErrorCode> {
  try {
    const result = await convertPptToPptx(bytes)
    const zip = await JSZip.loadAsync(result.bytes)
    for (const name of Object.keys(zip.files).filter((n) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(n),
    )) {
      const xml = await zip.files[name]!.async('string')
      expect(xml).not.toMatch(/NaN|Infinity/)
      for (const tag of xml.matchAll(/<(?:a:off|a:ext|a:chOff|a:chExt)\b[^>]*\/?>/g)) {
        for (const attr of tag[0].matchAll(/(?:x|y|cx|cy)="(-?\d+)"/g)) {
          const value = Number(attr[1])
          expect(Number.isFinite(value)).toBe(true)
          expect(Math.abs(value)).toBeLessThan(2 ** 53)
        }
      }
    }
    return 'ok'
  } catch (error) {
    expect(error).toBeInstanceOf(LegacyPptError)
    const code = (error as LegacyPptError).code
    expect(LEGACY_PPT_ERROR_CODES).toContain(code)
    return code
  }
}

/* ---------- byte surgery on a healthy deck ---------- */

const healthyDeck = (): Uint8Array => {
  const { document, currentUser } = buildPptStreams([
    {
      textSequences: [
        { kind: 4, text: 'Hello' },
        { kind: 4, text: 'World' },
      ],
    },
  ])
  return buildPptContainer(document, currentUser)
}

function mulberry32(seed: number): () => number {
  let a = seed >>> 0
  return () => {
    a = (a + 0x6d2b79f5) | 0
    let t = Math.imul(a ^ (a >>> 15), 1 | a)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

/** flip `count` pseudo-random bytes of a copy */
function flipped(bytes: Uint8Array, seed: number, count: number): Uint8Array {
  const out = new Uint8Array(bytes)
  const rand = mulberry32(seed)
  for (let i = 0; i < count; i++) {
    const at = Math.floor(rand() * out.length)
    out[at] = (out[at] ^ (1 << Math.floor(rand() * 8))) & 0xff
  }
  return out
}

/** the offsets of every record header whose type matches, scanning even offsets */
function recordHeaderOffsets(bytes: Uint8Array, type: number): number[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const hits: number[] = []
  for (let at = 0; at + 8 <= bytes.length; at += 2) {
    if (view.getUint16(at + 2, true) === type) hits.push(at)
  }
  return hits
}

/** lie about a record's declared length: the first header of `type` */
function withRecordLength(bytes: Uint8Array, type: number, newLength: number): Uint8Array {
  const out = new Uint8Array(bytes)
  const at = recordHeaderOffsets(out, type)[0]
  if (at !== undefined) new DataView(out.buffer).setUint32(at + 4, newLength >>> 0, true)
  return out
}

/* ---------- hostile drawing crafting (Escher by hand) ---------- */

const concatBytes = (parts: Uint8Array[]): Uint8Array => {
  const total = parts.reduce((sum, p) => sum + p.length, 0)
  const out = new Uint8Array(total)
  let at = 0
  for (const part of parts) {
    out.set(part, at)
    at += part.length
  }
  return out
}
const rec = (type: number, payload: Uint8Array, instance = 0, version = 0): Uint8Array => {
  const out = new Uint8Array(8 + payload.length)
  const view = new DataView(out.buffer)
  view.setUint16(0, ((instance & 0xfff) << 4) | (version & 0xf), true)
  view.setUint16(2, type, true)
  view.setUint32(4, payload.length, true)
  out.set(payload, 8)
  return out
}
const cont = (type: number, children: Uint8Array[]): Uint8Array =>
  rec(type, concatBytes(children), 0, 0xf)
const u16v = (value: number): Uint8Array => {
  const out = new Uint8Array(2)
  new DataView(out.buffer).setUint16(0, value, true)
  return out
}
const u32v = (value: number): Uint8Array => {
  const out = new Uint8Array(4)
  new DataView(out.buffer).setUint32(0, value >>> 0, true)
  return out
}
const i16v = (value: number): Uint8Array => {
  const out = new Uint8Array(2)
  new DataView(out.buffer).setInt16(0, value, true)
  return out
}
const utf16 = (text: string): Uint8Array => new Uint8Array(Buffer.from(text, 'utf16le'))
const rect32 = (r: [number, number, number, number]): Uint8Array => concatBytes(r.map(u32v))
const fspAtom = (shapeType: number, flags: number, spid: number): Uint8Array =>
  rec(0xf00a, concatBytes([u32v(spid), u32v(flags)]), shapeType, 2)
const clientAnchorAtom = (a: [number, number, number, number]): Uint8Array =>
  rec(0xf010, concatBytes(a.map(i16v)))
const childAnchorAtom = (r: [number, number, number, number]): Uint8Array => rec(0xf00f, rect32(r))
const spgrAtom = (r: [number, number, number, number]): Uint8Array => rec(0xf009, rect32(r), 0, 1)

/** a text shape; inside a group it carries the child anchor, outside its own */
const textboxLeaf = (spid: number, childAnchor?: [number, number, number, number]): Uint8Array =>
  cont(0xf004, [
    fspAtom(202, 0xa00, spid),
    ...(childAnchor ? [childAnchorAtom(childAnchor)] : [clientAnchorAtom([0, 0, 2000, 1000])]),
    rec(0xf00d, concatBytes([rec(0x0f9f, u32v(4)), rec(0x0fa0, utf16('x'))])),
  ])

/** a nested group level: self container (own + child anchor) wrapping the members */
const groupNode = (
  spid: number,
  members: Uint8Array[],
  bounds: [number, number, number, number] = [0, 0, 100, 100],
): Uint8Array =>
  cont(0xf003, [
    cont(0xf004, [
      spgrAtom(bounds),
      fspAtom(0, 0x201, spid),
      clientAnchorAtom([0, 0, 5760, 4320]),
      childAnchorAtom([0, 0, 100, 100]),
    ]),
    ...members,
  ])

const deckWithDrawing = (spgrChildren: Uint8Array[]): Uint8Array => {
  const spec: SlideSpec = {
    textSequences: [],
    rawDrawing: rec(
      0x03ee,
      concatBytes([
        rec(0x03ef, new Uint8Array(20)),
        cont(0x040c, [cont(0xf002, [cont(0xf003, spgrChildren)])]),
      ]),
      0,
      0xf,
    ),
  }
  const { document, currentUser } = buildPptStreams([spec])
  return buildPptContainer(document, currentUser)
}

/* ---------- corruption modes ---------- */

describe('hostile bytes: a malformed file fails coded or converts, never raw', () => {
  it('an empty and a tiny file are not containers', async () => {
    expect(await outcome(new Uint8Array(0))).toBe('not_cfb')
    expect(await outcome(new Uint8Array(16))).toBe('not_cfb')
  })

  it('a flipped CFB signature is not a container', async () => {
    const bytes = healthyDeck()
    bytes[0] ^= 0xff
    expect(await outcome(bytes)).toBe('not_cfb')
  })

  it('a file cut anywhere fails cleanly', async () => {
    const deck = healthyDeck()
    for (const keep of [1, 512, Math.floor(deck.length / 3), deck.length - 1]) {
      const code = await outcome(deck.slice(0, keep))
      expect(typeof code).toBe('string')
    }
  })

  it('a slide container claiming beyond the stream is a damaged slide', async () => {
    // the lie surfaces where the slide resolves, so the code is bad_slide —
    // truncated is for records that die inside the record walk itself
    expect(await outcome(withRecordLength(healthyDeck(), 0x03ee, 0x00ffffff))).toBe('bad_slide')
  })

  it('a zero-length slide container is skipped, the rest converts', async () => {
    expect(await outcome(withRecordLength(healthyDeck(), 0x03ee, 0))).toBe('ok')
  })

  it('seeded single-bit flips never surface a raw error', async () => {
    const deck = healthyDeck()
    for (let seed = 1; seed <= 24; seed++) {
      await outcome(flipped(deck, seed, 4))
    }
  })

  it('a dense flip storm across one deck stays coded', async () => {
    await outcome(flipped(healthyDeck(), 99, 64))
  })
})

describe('hostile shapes: the caps hold', () => {
  it('a group chain twice the descent cap terminates', async () => {
    // MAX_SPGR_DEPTH is 64 on the parse side and MAX_GROUP_DEPTH is 32 on the
    // mapping side; 128 levels must stop somewhere finite
    let node: Uint8Array = textboxLeaf(9000, [0, 0, 100, 100])
    for (let level = 0; level < 128; level++) {
      node = groupNode(5000 + level, [node])
    }
    expect(await cleanOutcome(deckWithDrawing([node]))).toBe('ok')
  })

  it('more shapes than the per-slide cap converts, capped and finite', async () => {
    // MAX_SHAPES_PER_SLIDE is 20 000; the deck carries 20 400 and the extra
    // 400 are dropped rather than walked forever
    const leaves = Array.from({ length: 20_400 }, (_, i) => textboxLeaf(3000 + i))
    expect(await cleanOutcome(deckWithDrawing(leaves))).toBe('ok')
  })

  it('more text sequences than the cap is a coded refusal', async () => {
    // MAX_TEXT_SEQUENCES is 100 000
    const leaves = Array.from({ length: 100_001 }, (_, i) => textboxLeaf(3000 + (i % 20_000)))
    const code = await outcome(deckWithDrawing(leaves))
    expect(code).toBe('limits_exceeded')
  })

  it('more slides declared than the cap is a coded refusal', async () => {
    // MAX_SLIDES is 4096; a document whose outline declares 4097 entries
    const listParts = Array.from({ length: 4097 }, (_, i) =>
      rec(0x03f3, concatBytes([u32v(i + 1), u32v(0), u32v(0), u32v(256 + i), u32v(0)])),
    )
    const doc = rec(
      0x03e8,
      concatBytes([
        rec(0x03e9, concatBytes([u32v(5760), u32v(4320), u32v(5760), u32v(4320)])),
        cont(0x0ff0, listParts),
      ]),
      0,
      0xf,
    )
    expect(await outcome(buildPptContainer(doc, new Uint8Array(0)))).toBe('limits_exceeded')
  })
})

describe('hostile drawing: degenerate and exploding geometry', () => {
  it('a group whose child space is one point drops its members, not the process', async () => {
    const group = groupNode(5000, [textboxLeaf(5001, [0, 0, 100, 100])], [0, 0, 0, 0])
    expect(await cleanOutcome(deckWithDrawing([group]))).toBe('ok')
  })

  it('a negative-width child space mirrors instead of exploding', async () => {
    const group = cont(0xf003, [
      cont(0xf004, [
        spgrAtom([1000, 1000, -1000, -1000]),
        fspAtom(0, 0x201, 5000),
        clientAnchorAtom([100, 100, 4000, 3000]),
        childAnchorAtom([-500, -500, 500, 500]),
      ]),
      textboxLeaf(5001, [-500, -500, 500, 500]),
    ])
    expect(await cleanOutcome(deckWithDrawing([group]))).toBe('ok')
  })

  it('anchors at the 16-bit extremes stay finite through the writer', async () => {
    const extreme = cont(0xf004, [
      fspAtom(202, 0xa00, 5002),
      clientAnchorAtom([-32768, -32768, 32767, 32767]),
      rec(0xf00d, concatBytes([rec(0x0f9f, u32v(4)), rec(0x0fa0, utf16('edge'))])),
    ])
    expect(await cleanOutcome(deckWithDrawing([extreme, textboxLeaf(5003)]))).toBe('ok')
  })

  it('an exploding group scale writes huge-but-finite offsets, never NaN', async () => {
    // child space of one master unit mapped onto the whole slide, member at
    // the far end of the child space: the scale is ~5760x
    const group = cont(0xf003, [
      cont(0xf004, [
        spgrAtom([0, 0, 1, 1]),
        fspAtom(0, 0x201, 5000),
        clientAnchorAtom([0, 0, 5760, 4320]),
        childAnchorAtom([0, 0, 32767, 32767]),
      ]),
      textboxLeaf(5001, [0, 0, 32767, 32767]),
    ])
    expect(await cleanOutcome(deckWithDrawing([group]))).toBe('ok')
  })

  it('a zero-width shape turned a quarter circle converts', async () => {
    const shape = cont(0xf004, [
      fspAtom(202, 0xa00, 5004),
      rec(0xf00b, concatBytes([u16v(4 | 0x8000), u32v(Math.round(90 * 65536) >>> 0)]), 1, 3),
      clientAnchorAtom([100, 100, 100, 3000]),
    ])
    expect(await cleanOutcome(deckWithDrawing([shape]))).toBe('ok')
  })

  it('a group flipped over a degenerate space terminates', async () => {
    const group = cont(0xf003, [
      cont(0xf004, [
        spgrAtom([0, 0, 0, 0]),
        fspAtom(0, 0x201 | 0x40, 5005), // flipH on a zero-size space
        clientAnchorAtom([0, 0, 5760, 4320]),
        childAnchorAtom([0, 0, 100, 100]),
      ]),
      textboxLeaf(5006, [0, 0, 100, 100]),
    ])
    expect(await cleanOutcome(deckWithDrawing([group]))).toBe('ok')
  })

  it('duplicate shape ids still produce unique output ids', async () => {
    const leaves = [textboxLeaf(700), textboxLeaf(700), textboxLeaf(700)]
    expect(await cleanOutcome(deckWithDrawing(leaves))).toBe('ok')
    const result = await convertPptToPptx(deckWithDrawing(leaves))
    const zip = await JSZip.loadAsync(result.bytes)
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    const ids = [...xml.matchAll(/id="(\d+)"/g)].map((m) => m[1])
    expect(new Set(ids).size).toBe(ids.length)
  })

  it('an unknown shape type passes through the caps, not the stack', async () => {
    const shape = cont(0xf004, [fspAtom(0xffff, 0xa00, 5007), clientAnchorAtom([0, 0, 1000, 1000])])
    expect(await cleanOutcome(deckWithDrawing([shape]))).toBe('ok')
  })
})

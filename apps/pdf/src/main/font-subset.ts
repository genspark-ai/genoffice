import { readFileSync } from 'node:fs'
import { hbSubsetWasmPath } from './wasm-path'

/** hb-subset.wasm exports we call (numbers are pointers into its linear memory) */
interface HbSubset {
  memory: WebAssembly.Memory
  malloc(size: number): number
  free(ptr: number): void
  hb_blob_create(data: number, len: number, mode: number, userData: number, destroy: number): number
  hb_blob_destroy(blob: number): void
  hb_blob_get_data(blob: number, len: number): number
  hb_blob_get_length(blob: number): number
  hb_face_create(blob: number, index: number): number
  hb_face_destroy(face: number): void
  hb_face_reference_blob(face: number): number
  hb_set_add(set: number, value: number): void
  hb_set_clear(set: number): void
  hb_set_invert(set: number): void
  hb_subset_input_create_or_fail(): number
  hb_subset_input_destroy(input: number): void
  hb_subset_input_set(input: number, setType: number): number
  hb_subset_input_unicode_set(input: number): number
  hb_subset_or_fail(face: number, input: number): number
}

const HB_MEMORY_MODE_WRITABLE = 2
const HB_SUBSET_SETS_LAYOUT_FEATURE_TAG = 6

let hbPromise: Promise<HbSubset> | null = null

function loadHb(): Promise<HbSubset> {
  hbPromise ??= (async () => {
    const wasmBytes = readFileSync(hbSubsetWasmPath())
    const { instance } = await WebAssembly.instantiate(wasmBytes)
    return instance.exports as unknown as HbSubset
  })()
  return hbPromise
}

/** Serialized: the wasm instance has one linear memory */
let chain: Promise<unknown> = Promise.resolve()

/** Subset a TrueType font to the glyphs needed for text (layout features kept).
    Same hb-subset calls as the subset-font npm package, without its fontverter/
    lodash/require.resolve baggage that breaks under a bundled main process. */
export function subsetTtf(font: Buffer, text: string): Promise<Buffer> {
  const run = chain.then(() => subsetTtfInner(font, text))
  chain = run.catch(() => undefined)
  return run
}

/** Read one CFF INDEX at `p`, returning item [start,end) pairs and the end position */
function readCffIndex(b: Buffer, p: number): { items: [number, number][]; end: number } {
  const count = b.readUInt16BE(p)
  if (count === 0) return { items: [], end: p + 2 }
  const offSize = b[p + 2]!
  const readOff = (i: number) => b.readUIntBE(p + 3 + i * offSize, offSize)
  const dataStart = p + 3 + (count + 1) * offSize - 1
  const items: [number, number][] = []
  for (let i = 0; i < count; i++) items.push([dataStart + readOff(i), dataStart + readOff(i + 1)])
  return { items, end: dataStart + readOff(count) }
}

/** Operator -> last operands of a CFF DICT */
function parseCffDict(b: Buffer, start: number, end: number): Map<number, number[]> {
  const ops = new Map<number, number[]>()
  let operands: number[] = []
  let i = start
  while (i < end) {
    const b0 = b[i]!
    if (b0 <= 21) {
      let op = b0
      if (b0 === 12) op = 1200 + b[++i]!
      ops.set(op, operands)
      operands = []
      i++
    } else if (b0 === 28) {
      operands.push(b.readInt16BE(i + 1))
      i += 3
    } else if (b0 === 29) {
      operands.push(b.readInt32BE(i + 1))
      i += 5
    } else if (b0 === 30) {
      // real number: nibbles until an 0xf terminator
      i++
      while (i < end && (b[i]! & 0x0f) !== 0x0f && b[i]! >> 4 !== 0x0f) i++
      i++
      operands.push(NaN)
    } else if (b0 >= 32 && b0 <= 246) {
      operands.push(b0 - 139)
      i++
    } else if (b0 >= 247 && b0 <= 250) {
      operands.push((b0 - 247) * 256 + b[i + 1]! + 108)
      i += 2
    } else if (b0 >= 251 && b0 <= 254) {
      operands.push(-(b0 - 251) * 256 - b[i + 1]! - 108)
      i += 2
    } else {
      i++
    }
  }
  return ops
}

/**
 * Rewrite a CID-keyed CFF's charset to the identity mapping, in a copy of the font.
 *
 * PDFium authors text by GID (Identity-H with CID=GID), but a CID-keyed CFF subset keeps
 * the ORIGINAL CIDs in its charset — conforming viewers (Acrobat/mupdf/poppler) resolve
 * CID→glyph through the charset, find nothing at CID=GID, and render blanks. An identity
 * charset (glyph i ⇔ CID i) makes both interpretations agree. The 5-byte format-2 range
 * always fits over the old charset in place (offsets elsewhere stay valid; the leftover
 * bytes become dead space, which CFF permits). Throws when the font is not rewritable —
 * callers fall back to another face. Non-CID CFFs and glyf-flavored fonts pass through.
 */
export function identityCffCharset(font: Buffer): Buffer {
  if (font.length < 12 || font.readUInt32BE(0) !== 0x4f54544f) return font // not OTTO
  const numTables = font.readUInt16BE(4)
  let cffOff = -1
  for (let t = 0; t < numTables; t++) {
    if (font.toString('latin1', 12 + 16 * t, 12 + 16 * t + 4) === 'CFF ') {
      cffOff = font.readUInt32BE(12 + 16 * t + 8)
      break
    }
  }
  if (cffOff < 0) return font
  const hdrSize = font[cffOff + 2]!
  const nameIndex = readCffIndex(font, cffOff + hdrSize)
  const topIndex = readCffIndex(font, nameIndex.end)
  const top = topIndex.items[0]
  if (!top) throw new Error('CFF top DICT missing')
  const ops = parseCffDict(font, top[0], top[1])
  if (!ops.has(1230)) return font // not CID-keyed: GID addressing already works
  const charsetOff = ops.get(15)?.[0]
  const charStringsOff = ops.get(17)?.[0]
  if (charsetOff === undefined || charsetOff <= 2 || charStringsOff === undefined)
    throw new Error('CID-keyed CFF without a custom charset')
  const nGlyphs = font.readUInt16BE(cffOff + charStringsOff)
  const cs = cffOff + charsetOff
  // The identity range must fit inside the old charset (in-place rewrite): size it first
  const fmt = font[cs]!
  let oldSize: number
  if (fmt === 0) {
    oldSize = 1 + 2 * (nGlyphs - 1)
  } else if (fmt === 1 || fmt === 2) {
    const step = fmt === 1 ? 3 : 4 // Range1: u16 first + u8 nLeft; Range2: u16 first + u16 nLeft
    let covered = 1
    let p = cs + 1
    while (covered < nGlyphs) {
      covered += (fmt === 1 ? font[p + 2]! : font.readUInt16BE(p + 2)) + 1
      p += step
    }
    oldSize = p - cs
  } else {
    throw new Error(`unknown CFF charset format ${fmt}`)
  }
  const out = Buffer.from(font)
  if (nGlyphs >= 2 && oldSize >= 5) {
    out[cs] = 2 // one range: glyphs 1..nGlyphs-1 become CIDs 1..nGlyphs-1
    out.writeUInt16BE(1, cs + 1)
    out.writeUInt16BE(nGlyphs - 2, cs + 3)
  } else if (nGlyphs >= 2 && oldSize === 4 && nGlyphs - 2 <= 0xff) {
    out[cs] = 1
    out.writeUInt16BE(1, cs + 1)
    out[cs + 3] = nGlyphs - 2
  } else if (nGlyphs !== 1) {
    throw new Error('CFF charset too small for an in-place identity rewrite')
  }
  return out
}

async function subsetTtfInner(font: Buffer, text: string): Promise<Buffer> {
  const hb = await loadHb()
  const input = hb.hb_subset_input_create_or_fail()
  if (!input) throw new Error('hb_subset_input_create_or_fail failed')

  const fontPtr = hb.malloc(font.byteLength)
  // Views are re-taken after every alloc: malloc may grow (and detach) the memory
  new Uint8Array(hb.memory.buffer).set(font, fontPtr)
  const blob = hb.hb_blob_create(fontPtr, font.byteLength, HB_MEMORY_MODE_WRITABLE, 0, 0)
  const face = hb.hb_face_create(blob, 0)
  hb.hb_blob_destroy(blob)

  let subset = 0
  let result = 0
  try {
    // Keep all layout features (equivalent of hb-subset --font-features='*')
    const features = hb.hb_subset_input_set(input, HB_SUBSET_SETS_LAYOUT_FEATURE_TAG)
    hb.hb_set_clear(features)
    hb.hb_set_invert(features)

    const unicodes = hb.hb_subset_input_unicode_set(input)
    for (const ch of text) hb.hb_set_add(unicodes, ch.codePointAt(0)!)

    subset = hb.hb_subset_or_fail(face, input)
    if (!subset) throw new Error('hb_subset_or_fail failed (corrupt font?)')

    result = hb.hb_face_reference_blob(subset)
    const len = hb.hb_blob_get_length(result)
    if (!len) throw new Error('hb-subset produced an empty font')
    const data = hb.hb_blob_get_data(result, 0)
    return Buffer.from(new Uint8Array(hb.memory.buffer).subarray(data, data + len))
  } finally {
    if (result) hb.hb_blob_destroy(result)
    if (subset) hb.hb_face_destroy(subset)
    hb.hb_face_destroy(face)
    hb.hb_subset_input_destroy(input)
    hb.free(fontPtr)
  }
}

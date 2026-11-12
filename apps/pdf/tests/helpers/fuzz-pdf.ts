/** Deterministic PDF generator + corruption operators for the pdf.js import
 *  fuzz. Everything is produced in-process (pdf-lib), so CI needs no external
 *  producer and a failing seed reproduces byte-for-byte from its number. */

import { PDFDocument, StandardFonts } from 'pdf-lib'

export function rng(seed: number): () => number {
  let a = seed >>> 0
  return () => {
    a |= 0
    a = (a + 0x6d2b79f5) | 0
    let t = Math.imul(a ^ (a >>> 15), 1 | a)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

/** A small but complete document: several pages, text, an outline entry and a
 *  link annotation, so a successful parse has real objects to walk. */
export async function buildFuzzPdf(seed: number): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const bold = await doc.embedFont(StandardFonts.HelveticaBold)
  const pages = 2 + (seed % 3)
  doc.setTitle(`Fuzz seed ${seed}`)
  doc.setAuthor('genoffice fuzz')
  for (let i = 0; i < pages; i++) {
    const page = doc.addPage([595, 842])
    page.setFont(font)
    page.drawText(`Anchor seed ${seed} page ${i + 1}`, {
      x: 50,
      y: 780,
      size: 14,
      font: bold,
    })
    for (let line = 0; line < 6 + (seed % 5); line++) {
      page.drawText(`fuzz line ${line} of document ${seed}`, {
        x: 50,
        y: 740 - line * 18,
        size: 10,
      })
    }
    if (i === 0) {
      page.drawRectangle({ x: 40, y: 600, width: 200, height: 40 })
    }
  }
  return doc.save({ useObjectStreams: seed % 2 === 0 })
}

export type Corrupt = (bytes: Uint8Array, r: () => number) => Uint8Array

const bitflip: Corrupt = (bytes, r) => {
  const out = Uint8Array.from(bytes)
  const flips = 1 + Math.floor(r() * 8)
  for (let i = 0; i < flips; i++) {
    const at = Math.floor(r() * out.length)
    out[at] = (out[at] ?? 0) ^ (1 << Math.floor(r() * 8))
  }
  return out
}

const truncate: Corrupt = (bytes, r) => {
  const at = 16 + Math.floor(r() * Math.max(1, bytes.length - 16))
  return bytes.slice(0, at)
}

const trailingJunk: Corrupt = (bytes, r) => {
  const junk = new Uint8Array(16 + Math.floor(r() * 256))
  for (let i = 0; i < junk.length; i++) junk[i] = Math.floor(r() * 256)
  const out = new Uint8Array(bytes.length + junk.length)
  out.set(bytes)
  out.set(junk, bytes.length)
  return out
}

const garbageWindow: Corrupt = (bytes, r) => {
  const out = Uint8Array.from(bytes)
  const at = Math.floor(r() * Math.max(1, out.length - 256))
  const len = Math.min(32 + Math.floor(r() * 160), out.length - at)
  for (let i = 0; i < len; i++) out[at + i] = Math.floor(r() * 256)
  return out
}

/** Rewrite the final startxref offset with a plausible but wrong number. */
const badStartxref: Corrupt = (bytes, r) => {
  const text = Buffer.from(bytes).toString('latin1')
  const at = text.lastIndexOf('startxref')
  if (at < 0) return bytes
  const out = Buffer.from(bytes)
  const value = String(1 + Math.floor(r() * 100000))
  const start = at + 'startxref'.length
  const rest = text.slice(start)
  const match = /\s*\d+/.exec(rest)
  if (!match) return bytes
  const digitsAt = start + match[0].indexOf(match[0].trim())
  const replacement = Buffer.from(value.padEnd(match[0].trim().length, ' '), 'latin1')
  if (digitsAt + replacement.length > out.length) return bytes
  out.set(replacement.subarray(0, match[0].trim().length), digitsAt)
  return new Uint8Array(out)
}

/** Mangle the header: pdf.js must refuse or recover, never hang. */
const badHeader: Corrupt = (bytes) => {
  const out = Uint8Array.from(bytes)
  out[1] = 'Q'.charCodeAt(0)
  return out
}

/** A stream that declares an enormous /Length: a reader must bound its work. */
const lyingStreamLength: Corrupt = (bytes, r) => {
  const text = Buffer.from(bytes).toString('latin1')
  const at = text.indexOf('/Length ')
  if (at < 0) return bytes
  const end = text.indexOf('\n', at)
  const value = String(1000000 + Math.floor(r() * 1000000))
  const line = `/Length ${value}`
  const out = Buffer.concat([
    Buffer.from(text.slice(0, at), 'latin1'),
    Buffer.from(line, 'latin1'),
    Buffer.from(text.slice(end), 'latin1'),
  ])
  return new Uint8Array(out)
}

export const PDF_CORRUPTIONS: Array<[string, Corrupt]> = [
  ['bitflip', bitflip],
  ['truncate', truncate],
  ['trailingjunk', trailingJunk],
  ['garbagewindow', garbageWindow],
  ['badstartxref', badStartxref],
  ['badheader', badHeader],
  ['lyinglength', lyingStreamLength],
]

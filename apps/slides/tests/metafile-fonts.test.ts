import { describe, expect, it } from 'vitest'
import { gzipSync } from 'node:zlib'
import { scanMetafileFonts } from '../src/main/metafile-fonts'

function u32(v: number): number[] {
  return [v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff, (v >>> 24) & 0xff]
}
function u16(v: number): number[] {
  return [v & 0xff, (v >> 8) & 0xff]
}
function utf16(s: string, chars: number): number[] {
  const out: number[] = []
  for (let i = 0; i < chars; i++) out.push(...u16(i < s.length ? s.charCodeAt(i) : 0))
  return out
}

function emfHeader(): number[] {
  const rec = new Array<number>(88).fill(0)
  rec.splice(0, 8, ...u32(1), ...u32(88))
  rec.splice(40, 4, ...u32(0x464d4520))
  return rec
}
function extCreateFontW(face: string, weight: number, italic: boolean): number[] {
  const lf = [
    ...u32(-24),
    ...u32(0),
    ...u32(0),
    ...u32(0),
    ...u32(weight),
    italic ? 1 : 0,
    0,
    0,
    0,
    0,
    0,
    0,
    0,
    ...utf16(face, 32),
  ]
  const body = [...u32(0), ...lf]
  while ((body.length + 8) % 4) body.push(0)
  return [...u32(82), ...u32(body.length + 8), ...body]
}
function emfPlusFontComment(face: string, styleFlags: number): number[] {
  const name = utf16(face, face.length)
  const font = [
    ...u32(0xdbc01002),
    ...u32(0x41400000),
    ...u32(0),
    ...u32(styleFlags),
    ...u32(0),
    ...u32(face.length),
    ...name,
  ]
  const rec = [
    ...u16(0x4008),
    ...u16((6 << 8) | 1),
    ...u32(12 + font.length),
    ...u32(font.length),
    ...font,
  ]
  const data = [...u32(0x2b464d45), ...rec]
  return [...u32(70), ...u32(12 + data.length), ...u32(data.length), ...data]
}

describe('scanMetafileFonts', () => {
  it('collects LOGFONTW facenames with weight/italic and EMF+ font objects', () => {
    const bytes = new Uint8Array([
      ...emfHeader(),
      ...extCreateFontW('Yu Gothic UI', 700, false),
      ...extCreateFontW('Meiryo UI', 400, true),
      ...extCreateFontW('Yu Gothic UI', 700, false),
      ...emfPlusFontComment('Segoe UI', 1),
    ])
    expect(scanMetafileFonts(bytes)).toEqual([
      { family: 'Yu Gothic UI', bold: true, italic: false },
      { family: 'Meiryo UI', bold: false, italic: true },
      { family: 'Segoe UI', bold: true, italic: false },
    ])
  })

  it('reads gzip-compressed payloads and WMF LOGFONT16 facenames', () => {
    const face = 'Arial Narrow'
    const lf16 = [
      ...u16(0xfff0),
      ...u16(0),
      ...u16(0),
      ...u16(0),
      ...u16(700),
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      ...[...face].map((c) => c.charCodeAt(0)),
      0,
    ]
    while (lf16.length % 2) lf16.push(0)
    const rec = [...u32((6 + lf16.length) / 2), ...u16(0x02fb), ...lf16]
    const header = [...u16(1), ...u16(9), ...u16(0x300), ...u32(0), ...u16(0), ...u32(0), ...u16(0)]
    const wmf = new Uint8Array([...header, ...rec])
    expect(scanMetafileFonts(wmf)).toEqual([{ family: 'Arial Narrow', bold: true, italic: false }])
    expect(scanMetafileFonts(new Uint8Array(gzipSync(wmf)))).toEqual([
      { family: 'Arial Narrow', bold: true, italic: false },
    ])
  })

  it('returns nothing for non-metafile bytes', () => {
    expect(scanMetafileFonts(new Uint8Array(64))).toEqual([])
  })
})

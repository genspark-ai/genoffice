import { describe, expect, it } from 'vitest'
import { latexToOmml, mathTokensOf, ommlToLatex, ommlToMathML } from '../src/math'
import {
  extractEquationNativeStream,
  isMathTypeProgId,
  mtefToOmml,
  oleBytesToOmml,
} from '../src/mtef'

const V5 = [5, 1, 0, 7, 0, 68, 83, 77, 84, 55, 0, 1]
const V3 = [3, 1, 1, 3, 0]
const NULL = [1, 1]
const EOC = 0xfffffffe
const FREE = 0xffffffff

const word = (value: number) => [value & 255, value >> 8]
const cString = (value: string) => [...value].map((ch) => ch.charCodeAt(0)).concat(0)
const ch = (glyph: string, face = 3) => [2, 0, 128 + face, ...word(glyph.charCodeAt(0))]
const line = (...records: number[][]) => [1, 0, ...records.flat(), 0]
const template = (selector: number, variation: number, ...records: number[][]) => [
  3,
  0,
  selector,
  ...(variation >= 128 ? [0x80 | (variation & 127), variation >> 8] : [variation]),
  0,
  ...records.flat(),
  0,
]
const equation = (...records: number[][]) => Uint8Array.from([...V5, ...line(...records), 0])
const legacyChar = (glyph: string, face = 3) => ch(glyph, face).filter((_, index) => index !== 1)
const legacyLine = (...records: number[][]) => [1, ...records.flat(), 0]
const legacyTemplate = (selector: number, variation: number, ...records: number[][]) => [
  3,
  selector,
  variation,
  0,
  ...records.flat(),
  0,
]

function fixtureOle(
  mtef: Uint8Array,
  { version = 3, mini = true, difat = false, headerSize = 28 } = {},
) {
  const sectorSize = version === 3 ? 512 : 4096
  const nativeSize = Math.max(mini ? 0 : 4096, headerSize + mtef.length)
  const miniStreamSize = Math.ceil(nativeSize / 64) * 64
  const dataSectors = Math.ceil(nativeSize / sectorSize)
  const dataId = mini ? 2 : 1
  const fatId = dataId + dataSectors
  const fatCount = difat ? 110 : 1
  const difatId = fatId + fatCount
  const bytes = new Uint8Array((difatId + Number(difat) + 1) * sectorSize)
  const data = new DataView(bytes.buffer)
  const u16 = (offset: number, value: number) => data.setUint16(offset, value, true)
  const u32 = (offset: number, value: number) => data.setUint32(offset, value, true)
  const sector = (id: number) => (id + 1) * sectorSize
  const directory = sector(0)
  const fat = sector(fatId)
  const miniFat = sector(1)
  bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
  u16(24, 0x3e)
  u16(26, version)
  u16(28, 0xfffe)
  u16(30, version === 3 ? 9 : 12)
  u16(32, 6)
  u32(40, version === 4 ? 1 : 0)
  u32(44, fatCount)
  u32(48, 0)
  u32(56, 4096)
  u32(60, mini ? 1 : EOC)
  u32(64, Number(mini))
  u32(68, difat ? difatId : EOC)
  u32(72, Number(difat))
  for (let i = 0; i < 109; i++) u32(76 + 4 * i, i < fatCount ? fatId + i : FREE)
  if (difat) {
    const offset = sector(difatId)
    bytes.fill(255, offset, offset + sectorSize)
    u32(offset, fatId + 109)
    u32(offset + sectorSize - 4, EOC)
  }
  bytes.fill(255, fat, fat + fatCount * sectorSize)
  u32(fat, EOC)
  for (let i = 0; i < fatCount; i++) u32(fat + (fatId + i) * 4, 0xfffffffd)
  if (difat) u32(fat + difatId * 4, 0xfffffffc)
  for (let i = 0; i < dataSectors; i++)
    u32(fat + (dataId + i) * 4, i + 1 === dataSectors ? EOC : dataId + i + 1)

  function directoryEntry(
    offset: number,
    name: string,
    type: number,
    child: number,
    start: number,
    size: number,
  ) {
    for (let i = 0; i < name.length; i++) u16(offset + i * 2, name.charCodeAt(i))
    u16(offset + 64, (name.length + 1) * 2)
    bytes.set([type, 1], offset + 66)
    u32(offset + 68, FREE)
    u32(offset + 72, FREE)
    u32(offset + 76, child)
    u32(offset + 116, start)
    u32(offset + 120, size)
  }
  const lastMini = miniStreamSize / 64 - 1
  directoryEntry(directory, 'Root Entry', 5, 1, mini ? dataId : EOC, mini ? miniStreamSize : 0)
  directoryEntry(directory + 128, 'Equation Native', 2, FREE, mini ? lastMini : dataId, nativeSize)

  const native = new Uint8Array(nativeSize)
  const header = new DataView(native.buffer)
  header.setUint16(0, headerSize, true)
  header.setUint32(2, 0x20000, true)
  header.setUint32(8, mtef.length, true)
  native.set(mtef, headerSize)
  const nativeOffset = sector(dataId) + (mini ? lastMini * 64 : 0)
  if (mini) {
    u32(fat + 4, EOC)
    bytes.fill(255, miniFat, miniFat + sectorSize)
    for (let i = lastMini; i >= 0; i--) {
      u32(miniFat + i * 4, i === 0 ? EOC : i - 1)
      const source = (lastMini - i) * 64
      bytes.set(native.subarray(source, source + 64), sector(dataId) + i * 64)
    }
  } else {
    bytes.set(native, nativeOffset)
  }
  return { bytes, directory, fat, miniFat, nativeOffset, dataId, lastMini, u16, u32 }
}

function decode(payload: Uint8Array): string {
  const decoded = mtefToOmml(payload)
  expect(decoded).toMatch(/^<m:oMath>.*<\/m:oMath>$/s)
  return decoded!
}

describe('Equation Native OLE stream', () => {
  it('recognizes MathType and Equation Editor ProgIDs', () => {
    for (const id of [
      'equation.dsmt4',
      'Equation.DSMT3',
      'Equation.DSMT7',
      'Equation.3',
      'Equation.2',
      'Equation',
    ]) {
      expect(isMathTypeProgId(id)).toBe(true)
    }
    for (const id of [undefined, 'Equation.30', 'Word.Document', 'Equation.DSMT9', '']) {
      expect(isMathTypeProgId(id)).toBe(false)
    }
  })

  it.each([
    { version: 3, mini: true },
    { version: 4, mini: true },
    { version: 3, mini: false },
    { version: 4, mini: false },
    { version: 3, mini: false, difat: true, headerSize: 32 },
  ])('extracts a bounded MTEF slice from a CFB $version stream (MiniFAT $mini)', (options) => {
    const mtef = equation(...Array.from({ length: 100 }, () => ch('x')))
    const { bytes } = fixtureOle(mtef, options)
    const withOffset = new Uint8Array(bytes.length + 12)
    withOffset.set(bytes, 7)
    expect(extractEquationNativeStream(withOffset.subarray(7, 7 + bytes.length))).toEqual(mtef)
    expect(oleBytesToOmml(bytes)).toBe(mtefToOmml(mtef))
  })

  it('rejects truncated CFBs, invalid headers, and missing streams', () => {
    const payload = equation(ch('x'))
    const file = fixtureOle(payload)
    expect(extractEquationNativeStream(new Uint8Array(0))).toBeNull()
    expect(extractEquationNativeStream(file.bytes.subarray(0, file.bytes.length - 1))).toBeNull()
    file.u16(file.nativeOffset, 27)
    expect(extractEquationNativeStream(file.bytes)).toBeNull()
    file.u16(file.nativeOffset, 28)
    file.u32(file.nativeOffset + 8, 0xffffffff)
    expect(extractEquationNativeStream(file.bytes)).toBeNull()
    file.u32(file.nativeOffset + 8, payload.length)
    file.u32(file.nativeOffset + 2, 0x30000)
    expect(extractEquationNativeStream(file.bytes)).toBeNull()
    file.u32(file.nativeOffset + 2, 0x20000)
    file.u16(file.directory + 128, 'X'.charCodeAt(0))
    expect(extractEquationNativeStream(file.bytes)).toBeNull()
  })

  it('rejects cycles in regular FAT, MiniFAT and DIFAT chains', () => {
    const payload = equation(...Array.from({ length: 25 }, () => ch('x')))
    const major = fixtureOle(payload, { mini: false })
    major.u32(major.fat + major.dataId * 4, major.dataId)
    expect(extractEquationNativeStream(major.bytes)).toBeNull()
    const minor = fixtureOle(payload)
    minor.u32(minor.miniFat + minor.lastMini * 4, minor.lastMini)
    expect(extractEquationNativeStream(minor.bytes)).toBeNull()
    const difat = fixtureOle(payload, { mini: false, difat: true })
    const difatId = new DataView(difat.bytes.buffer).getUint32(68, true)
    difat.u32((difatId + 2) * 512 - 4, difatId)
    expect(extractEquationNativeStream(difat.bytes)).toBeNull()
  })
})

describe('MTEF records and MathType styles', () => {
  it('handles v5 metadata, rulers, nudge offsets, and future extension lengths', () => {
    const future = [102, 255, ...word(256), ...new Uint8Array(256).fill(0xab)]
    const header = [
      21,
      ...cString('Unicode'),
      17,
      5,
      ...cString('Test'),
      19,
      4,
      ...word(0),
      ...word(10),
      ...word(0),
      ...cString('Ink'),
      18,
      1,
    ]
    const dimensions = [20, 0, 1, 0x21, 0xf0, 2, 0x41, 0x0f, 0x2f, 1, 1, 0]
    const sizes = [
      9,
      101,
      ...word(240),
      9,
      100,
      1,
      ...word(24),
      9,
      2,
      128,
      10,
      11,
      12,
      13,
      14,
      15,
      16,
    ]
    const nudged = [2, 8, 128, 128, 0, 0, 0, 0, 131, ...word(0x78)]
    const payload = Uint8Array.from([
      ...V5,
      ...future,
      ...header,
      ...dimensions,
      ...sizes,
      1,
      6,
      ...word(100),
      7,
      1,
      2,
      ...word(120),
      ...nudged,
      ...ch('<', 6),
      ...ch('&', 6),
      ...ch('>', 6),
      0,
      0,
    ])
    const omml = decode(payload)
    expect(mathTokensOf(omml)).toEqual(['x', '<', '&', '>'])
    expect(omml).toContain('&lt;')
    expect(ommlToMathML(omml)).toContain('<mo>&amp;</mo>')
  })

  it('uses v5 character styles, MTCode and font-specific positions', () => {
    const definitions = [
      19,
      ...cString('EuclidMath1'),
      17,
      5,
      ...cString('Euclid Math One'),
      8,
      1,
      0,
      17,
      1,
      ...cString('Times New Roman'),
      8,
      2,
      1,
      8,
      2,
      3,
      17,
      1,
      ...cString('Blackboard'),
      8,
      3,
      0,
      17,
      1,
      ...cString('Fraktur'),
      8,
      4,
      0,
    ]
    const payload = Uint8Array.from([
      ...V5,
      ...definitions,
      ...line(
        ch('N', 1),
        ch('w', -2),
        [2, 0, 255, ...word(32765), ...word(0x03b7)],
        [2, 4, 127, ...word(0xf10a), 0x4b],
        ch('\u2112', -1),
        ch('R', -4),
        ch('C', -5),
        [2, 0x24, 132, 0x61],
        [2, 0x24, 134, 0xa3],
        [2, 0x30, 139, ...word(0xf04c)],
        [2, 4, 139, ...word(0x2113), 0x6c],
        ch('\uef04', 24),
        ch('\ue901', 6),
      ),
      0,
    ])
    const omml = decode(payload)
    expect(mathTokensOf(omml)).toEqual([
      'N',
      'w',
      '\u03b7',
      'K',
      'L',
      'R',
      'C',
      '\u03b1',
      '\u2264',
      '\u22ef',
      '\u2113',
      '\u2004',
      '\u2a72',
    ])
    expect(omml).toContain('<m:sty m:val="p"/>')
    expect(omml).toContain('<m:sty m:val="bi"/>')
    expect(omml).toContain('<m:scr m:val="script"/>')
    const mml = ommlToMathML(omml)
    for (const glyph of [
      '\u{1d430}',
      '\u{1d4a6}',
      '\u2112',
      '\u211d',
      '\u212d',
      '\u03b1',
      '\u2264',
      '\u22ef',
    ])
      expect(mml).toContain(glyph)
    const latex = ommlToLatex(omml)
    expect(latex).toContain('\\mathbf{w}')
    expect(latex).toContain('\\mathcal{K}')
    expect(latex).toContain('\\mathbb{R}')
    expect(latex).toContain('\\mathfrak{C}')
  })

  it('uses v3 packed tags, FONT records, and legacy selectors', () => {
    const payload = Uint8Array.from([
      ...V3,
      8,
      1,
      2,
      ...cString('Times New Roman'),
      10,
      ...legacyLine(
        legacyChar('x', -1),
        legacyTemplate(15, 1, legacyLine(legacyChar('i')), [0x11]),
        legacyChar('=', 6),
        legacyTemplate(14, 0, legacyLine(legacyChar('a')), legacyLine(legacyChar('b'))),
        [0x22, 131, ...word(0x79), 6, 9, 0],
      ),
      0,
    ])
    const omml = decode(payload)
    expect(ommlToLatex(omml)).toBe('{\\mathbf{x}}_{i}=\\frac{a}{b}\\hat{y}')
    expect(ommlToMathML(omml)).toContain('<mfrac>')
  })

  it('reads padded v3 embellishments and font-specific fence pieces', () => {
    const payload = Uint8Array.from([
      ...V3,
      ...legacyLine(
        legacyTemplate(
          7,
          0,
          legacyLine([0x32, 131, ...word(0x78), 6, 8, 0, 0]),
          legacyChar('\uf8ee', 22),
          legacyChar('\uf8f9', 22),
        ),
      ),
      0,
    ])
    const omml = decode(payload)
    expect(ommlToLatex(omml)).toBe('\\left\\lceil \\tilde{x} \\right\\rceil')
    expect(ommlToMathML(omml)).toContain('<mo stretchy="true">\u2308</mo>')
  })
})

describe('MTEF equation structures', () => {
  it('attaches sibling scripts, merging complementary scripts on the same base', () => {
    const low = line(ch('i'))
    const high = line(ch('2', 8))
    const omml = decode(
      equation(
        ch('a'),
        template(27, 0, low, NULL),
        ch('+', 6),
        ch('b'),
        template(28, 0, NULL, high),
        ch('+', 6),
        ch('c'),
        template(29, 0, low, high),
        ch('+', 6),
        ch('d'),
        template(28, 0, NULL, high),
        template(27, 0, low, NULL),
      ),
    )
    expect(ommlToLatex(omml)).toBe('{a}_{i}+{b}^{2}+{c}_{i}^{2}+{d}_{i}^{2}')
    expect(ommlToMathML(omml).match(/<msubsup>/g)).toHaveLength(2)
    const pre = decode(equation(template(27, 1, low, NULL), template(28, 1, NULL, high), ch('T')))
    expect(pre).toContain('<m:sPre><m:sub>')
    expect(ommlToMathML(pre)).toContain('<mmultiscripts><mrow><mi>T</mi></mrow><mprescripts/>')
  })

  it('converts roots, fractions, and fences and round trips through LaTeX', () => {
    const radical = template(10, 1, line(ch('x')), line(ch('3', 8)))
    const fraction = template(11, 0, line(radical), line(ch('y')))
    const omml = decode(equation(template(1, 3, line(fraction), ch('(', 6), ch(')', 6))))
    const latex = ommlToLatex(omml)
    expect(latex).toBe('\\left( \\frac{\\sqrt[3]{x}}{y} \\right)')
    expect(ommlToMathML(`<m:oMath>${latexToOmml(latex!)}</m:oMath>`)).toBe(ommlToMathML(omml))
    const half = decode(
      equation(template(2, 1, line(template(10, 0, line(ch('x')), NULL)), ch('{', 6))),
    )
    expect(half).toContain('<m:endChr m:val=""/>')
    expect(ommlToMathML(half)).toContain('<msqrt>')
  })

  it('preserves limits and glyphs for summation and integral styles', () => {
    const x = line(ch('x'))
    const low = line(ch('0', 8))
    const high = line(ch('1', 8))
    const omml = decode(
      equation(
        template(16, 0x70, x, low, high, ch('\u2211', 6)),
        template(15, 0x131, x, low, high, ch('\u222b', 6)),
        template(15, 0x12, x, low, NULL, ch('\u222b', 6), ch('\u222b', 6)),
        template(17, 0x50, x, low, NULL),
        template(20, 0x60, x, NULL, high),
        template(21, 0x10, x, low, NULL, ch('\u2295', 6)),
      ),
    )
    const mml = ommlToMathML(omml)
    expect(mml).toContain(
      '<munderover><mo stretchy="false">\u2211</mo><mrow><mn>0</mn></mrow><mrow><mn>1</mn></mrow>',
    )
    expect(mml).toContain(
      '<msubsup><mo stretchy="false">\u222b</mo><mrow><mn>0</mn></mrow><mrow><mn>1</mn></mrow>',
    )
    expect(mml).toContain('<msub><mo stretchy="false">\u222c</mo>')
    expect(mml).toContain('<munder><mo stretchy="false">\u220f</mo>')
    expect(mml).toContain('<mover><mo stretchy="false">\u22c2</mo>')
    expect(mml).toContain('<msub><mo stretchy="false">\u2295</mo>')
  })

  it('renders matrices, piles, and decorated expressions', () => {
    const a = line(ch('a'))
    const b = line(ch('b'))
    const matrix = [5, 0, 0, 1, 0, 2, 2, 0, 0, ...a, ...b, ...b, ...a, 0]
    const pile = [4, 2, 2, 1, 7, 1, 2, 0, 0, ...a, ...b, 0]
    const embell = [2, 1, 131, ...word(120), 6, 0, 2, 6, 8, 1, 1, 5, 0]
    const omml = decode(
      equation(
        matrix,
        pile,
        embell,
        template(31, 2, a),
        template(32, 0, a),
        template(33, 0, a),
        template(34, 0, a),
        template(12, 0, a),
        template(13, 0, a),
        template(24, 1, a, b),
        template(37, 0x1e, b),
        template(14, 0x2c, a, b, ch('\u2192', 6)),
        template(23, 0x30, line(ch('l', 2), ch('i', 2), ch('m', 2)), a, b),
      ),
    )
    const mml = ommlToMathML(omml)
    expect(mml.match(/<mtr>/g)).toHaveLength(4)
    expect(mml).toContain('text-align:left')
    expect(mml).toContain('<mover accent="true">')
    expect(mml).toContain('\u2032')
    expect(mml).toContain('border-top:0.06em solid')
    for (const tag of [
      'm:eqArr',
      'm:bar',
      'm:groupChr',
      'm:limLow',
      'm:limUpp',
      'm:borderBox',
      'm:sSup',
    ])
      expect(omml).toContain(`<${tag}>`)
  })

  it('returns null for missing data, mismatched structures and excessive nesting', () => {
    const good = equation(ch('x'), template(27, 0, line(ch('i')), NULL))
    for (let size = 0; size < good.length; size++)
      expect(mtefToOmml(good.subarray(0, size))).toBeNull()
    for (const bad of [
      equation(),
      equation([2, 0x20, 131]),
      equation(ch('\uffff')),
      equation(ch('\ue000')),
      equation(template(11, 0, line(ch('x')))),
      equation(template(255, 0, line(ch('x')))),
      equation(template(27, 0, line(ch('i')))),
      equation(template(28, 1, line(ch('i')))),
      equation([5, 0, 0, 0, 0, 2, 2, 0, 0, ...line(ch('x')), 0]),
      Uint8Array.from([...V5, 100, 255, 255, 255, 0]),
      Uint8Array.from([...V5, ...new Uint8Array(20_001).fill(10), ...line(ch('x')), 0]),
      Uint8Array.from([
        ...V5,
        ...Array.from({ length: 34 }, () => [1, 0]).flat(),
        ...ch('x'),
        ...new Uint8Array(35),
      ]),
    ])
      expect(mtefToOmml(bad)).toBeNull()
    expect(oleBytesToOmml(fixtureOle(good.subarray(0, good.length - 1)).bytes)).toBeNull()
  })
})

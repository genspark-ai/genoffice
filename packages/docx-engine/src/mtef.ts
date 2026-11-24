const MAX_OLE_BYTES = 32 * 1024 * 1024
const MAX_MTEF_BYTES = 1024 * 1024
const MAX_RECORDS = 20_000
const MAX_DEPTH = 32
const FREE = 0xffffffff
const END = 0xfffffffe
const FAT = 0xfffffffd
const DIFAT = 0xfffffffc
const MATH_TYPE_PROG_IDS = new Set([
  'equation',
  'equation.2',
  'equation.3',
  'equation.dsmt3',
  'equation.dsmt4',
  'equation.dsmt7',
])

class InvalidEquation extends Error {}

function requireValid(condition: unknown): asserts condition {
  if (!condition) throw new InvalidEquation()
}

function view(bytes: Uint8Array): DataView {
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
}

function hasInvalidXmlChar(text: string): boolean {
  for (const glyph of text) {
    const code = glyph.codePointAt(0)!
    if (
      (code <= 0x1f && code !== 0x09 && code !== 0x0a && code !== 0x0d) ||
      (code >= 0xd800 && code <= 0xdfff) ||
      code === 0xfffe ||
      code === 0xffff
    ) {
      return true
    }
  }
  return false
}

export function isMathTypeProgId(progId: string | undefined): boolean {
  return progId !== undefined && MATH_TYPE_PROG_IDS.has(progId.toLowerCase())
}

function equationStream(bytes: Uint8Array): Uint8Array {
  // Sector and directory layout per MS-CFB sections 2.2 through 2.6.
  requireValid(bytes.length >= 512 && bytes.length <= MAX_OLE_BYTES)
  const data = view(bytes)
  const u16 = (offset: number) => data.getUint16(offset, true)
  const u32 = (offset: number) => data.getUint32(offset, true)
  requireValid(u32(0) === 0xe011cfd0 && u32(4) === 0xe11ab1a1)
  const major = u16(26)
  const shift = u16(30)
  const dirCount = u32(40)
  const cutoff = u32(56)
  requireValid(
    (major === 3 && shift === 9 && dirCount === 0) || (major === 4 && shift === 12 && dirCount > 0),
  )
  requireValid(u16(28) === 0xfffe && u16(32) === 6 && cutoff === 4096)
  const sectorSize = 1 << shift
  const sectorCount = bytes.length / sectorSize - 1
  const entriesPerSector = sectorSize / 4
  requireValid(Number.isInteger(sectorCount) && sectorCount >= 2)
  const offsetOf = (id: number) => {
    requireValid(id < sectorCount)
    return (id + 1) * sectorSize
  }

  const fatCount = u32(44)
  const difatCount = u32(72)
  requireValid(
    fatCount > 0 &&
      fatCount + difatCount <= sectorCount &&
      fatCount * entriesPerSector >= sectorCount,
  )
  const fatOffsets: number[] = []
  const fatIds = new Set<number>()
  const difatIds = new Set<number>()
  const addFat = (id: number) => {
    if (id === FREE) return
    requireValid(fatOffsets.length < fatCount && !fatIds.has(id) && !difatIds.has(id))
    fatOffsets.push(offsetOf(id))
    fatIds.add(id)
  }
  for (let i = 0; i < 109; i++) addFat(u32(76 + i * 4))
  let difat = u32(68)
  for (let i = 0; i < difatCount; i++) {
    requireValid(!difatIds.has(difat) && !fatIds.has(difat))
    const offset = offsetOf(difat)
    difatIds.add(difat)
    for (let j = 0; j < entriesPerSector - 1; j++) addFat(u32(offset + j * 4))
    difat = u32(offset + sectorSize - 4)
  }
  requireValid(difat === END && fatOffsets.length === fatCount)
  const fatNext = (id: number) =>
    u32(fatOffsets[Math.floor(id / entriesPerSector)] + (id % entriesPerSector) * 4)
  for (const id of fatIds) requireValid(fatNext(id) === FAT)
  for (const id of difatIds) requireValid(fatNext(id) === DIFAT)

  function chain(
    start: number,
    next: (id: number) => number,
    capacity: number,
    count?: number,
  ): number[] {
    requireValid(count === undefined || count <= capacity)
    const ids: number[] = []
    const seen = new Set<number>()
    for (let id = start; id !== END; id = next(id)) {
      requireValid(id < capacity && !seen.has(id) && (count === undefined || ids.length < count))
      seen.add(id)
      ids.push(id)
    }
    requireValid(count === undefined || ids.length === count)
    return ids
  }

  const allocated = new Set<number>()
  function regularChain(start: number, count?: number): number[] {
    const ids = chain(start, fatNext, sectorCount, count)
    return ids.map((id) => {
      requireValid(!allocated.has(id) && !fatIds.has(id) && !difatIds.has(id))
      allocated.add(id)
      return offsetOf(id)
    })
  }

  const directory = regularChain(u32(48), major === 4 ? dirCount : undefined)
  requireValid(directory.length > 0 && bytes[directory[0] + 66] === 5)
  const root = directory[0]
  const name = 'Equation Native'
  let native: number | undefined
  for (const sector of directory) {
    for (let entry = sector; entry < sector + sectorSize; entry += 128) {
      if (
        bytes[entry + 66] !== 2 ||
        u16(entry + 64) !== (name.length + 1) * 2 ||
        u16(entry + name.length * 2) !== 0
      )
        continue
      let entryName = ''
      for (let j = 0; j < name.length; j++) entryName += String.fromCharCode(u16(entry + j * 2))
      if (entryName.toLowerCase() !== name.toLowerCase()) continue
      requireValid(native === undefined)
      native = entry
    }
  }
  requireValid(native !== undefined)
  const streamSize = (entry: number) => {
    requireValid(major === 3 || u32(entry + 124) === 0)
    return u32(entry + 120)
  }
  const size = streamSize(native)
  const start = u32(native + 116)
  requireValid(size >= 28 && size <= bytes.length && size <= MAX_MTEF_BYTES + 0xffff)
  const unit = size < cutoff ? 64 : sectorSize
  const count = Math.ceil(size / unit)
  let offsets: number[]
  if (size < cutoff) {
    const miniFat = regularChain(u32(60), u32(64))
    const rootSize = streamSize(root)
    requireValid(miniFat.length > 0 && rootSize <= bytes.length && rootSize % 64 === 0)
    const miniStream = regularChain(u32(root + 116), Math.ceil(rootSize / sectorSize))
    const miniNext = (id: number) =>
      u32(miniFat[Math.floor(id / entriesPerSector)] + (id % entriesPerSector) * 4)
    offsets = chain(
      start,
      miniNext,
      Math.min(rootSize / unit, miniFat.length * entriesPerSector),
      count,
    ).map((id) => miniStream[Math.floor((id * unit) / sectorSize)] + ((id * unit) % sectorSize))
  } else {
    offsets = regularChain(start, count)
  }
  const stream = new Uint8Array(size)
  offsets.forEach((offset, index) => {
    const target = index * unit
    stream.set(bytes.subarray(offset, offset + Math.min(unit, size - target)), target)
  })
  return stream
}

export function extractEquationNativeStream(oleBytes: Uint8Array): Uint8Array | null {
  try {
    const stream = equationStream(oleBytes)
    const header = view(stream)
    const headerSize = header.getUint16(0, true)
    const objectSize = header.getUint32(8, true)
    requireValid(headerSize >= 28 && header.getUint32(2, true) === 0x20000)
    requireValid(
      objectSize > 0 && objectSize <= MAX_MTEF_BYTES && headerSize + objectSize <= stream.length,
    )
    return stream.subarray(headerSize, headerSize + objectSize)
  } catch (error) {
    if (error instanceof InvalidEquation) return null
    throw error
  }
}

export function oleBytesToOmml(oleBytes: Uint8Array): string | null {
  const mtef = extractEquationNativeStream(oleBytes)
  return mtef ? mtefToOmml(mtef) : null
}

type ScriptFont = 'script' | 'fraktur' | 'double-struck' | 'sans-serif' | 'monospace'
type Style = 'p' | 'b' | 'i' | 'bi'
type Typeface = { name: string; encoding: string; flags: number }

const SYMBOL_UPPER = 'ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ'
const SYMBOL_LOWER = 'αβχδεϕγηιφκλμνοπθρστυϖωξψζ'
const SYMBOL_ANSI: Readonly<Record<number, string>> = {
  0x22: '∀',
  0x24: '∃',
  0x27: '∋',
  0x2a: '∗',
  0x2d: '−',
  0x40: '≅',
  0x5c: '∴',
  0x5e: '⊥',
  0x7e: '∼',
  0xa0: '€',
  0xa1: 'ϒ',
  0xa2: '′',
  0xa3: '≤',
  0xa4: '⁄',
  0xa5: '∞',
  0xa6: 'ƒ',
  0xa7: '♣',
  0xa8: '♦',
  0xa9: '♥',
  0xaa: '♠',
  0xab: '↔',
  0xac: '←',
  0xad: '↑',
  0xae: '→',
  0xaf: '↓',
  0xb0: '°',
  0xb1: '±',
  0xb2: '″',
  0xb3: '≥',
  0xb4: '×',
  0xb5: '∝',
  0xb6: '∂',
  0xb7: '•',
  0xb8: '÷',
  0xb9: '≠',
  0xba: '≡',
  0xbb: '≈',
  0xbc: '…',
  0xbd: '⏐',
  0xbe: '⎯',
  0xbf: '↵',
  0xc0: 'ℵ',
  0xc1: 'ℑ',
  0xc2: 'ℜ',
  0xc3: '℘',
  0xc4: '⊗',
  0xc5: '⊕',
  0xc6: '∅',
  0xc7: '∩',
  0xc8: '∪',
  0xc9: '⊃',
  0xca: '⊇',
  0xcb: '⊄',
  0xcc: '⊂',
  0xcd: '⊆',
  0xce: '∈',
  0xcf: '∉',
  0xd0: '∠',
  0xd1: '∇',
  0xd2: '®',
  0xd3: '©',
  0xd4: '™',
  0xd5: '∏',
  0xd6: '√',
  0xd7: '⋅',
  0xd8: '¬',
  0xd9: '∧',
  0xda: '∨',
  0xdb: '⇔',
  0xdc: '⇐',
  0xdd: '⇑',
  0xde: '⇒',
  0xdf: '⇓',
  0xe0: '◊',
  0xe1: '⟨',
  0xe2: '®',
  0xe3: '©',
  0xe4: '™',
  0xe5: '∑',
  0xe6: '⎛',
  0xe7: '⎜',
  0xe8: '⎝',
  0xe9: '⎡',
  0xea: '⎢',
  0xeb: '⎣',
  0xec: '⎧',
  0xed: '⎨',
  0xee: '⎩',
  0xef: '⎪',
  0xf1: '⟩',
  0xf2: '∫',
  0xf3: '⌠',
  0xf4: '⎮',
  0xf5: '⌡',
  0xf6: '⎞',
  0xf7: '⎟',
  0xf8: '⎠',
  0xf9: '⎤',
  0xfa: '⎥',
  0xfb: '⎦',
  0xfc: '⎫',
  0xfd: '⎬',
  0xfe: '⎭',
}

const MT_EXTRA: Readonly<Record<number, string>> = {
  0x20: ' ',
  0x2c: '\u200a',
  0x3a: '∼',
  0x3b: '≃',
  0x3c: '◁',
  0x3d: '≪',
  0x3e: '▷',
  0x3f: '≫',
  0x40: '≜',
  0x41: '≙',
  0x42: '≐',
  0x43: '∐',
  0x44: 'ƛ',
  0x45: '⨽',
  0x46: '⨼',
  0x47: '⌐',
  0x48: '￢',
  0x49: '⋂',
  0x4a: '℧',
  0x4b: '…',
  0x4c: '⋯',
  0x4d: '⋮',
  0x4e: '⋰',
  0x4f: '⋱',
  0x50: '∥',
  0x51: '∵',
  0x52: '∡',
  0x53: '∢',
  0x55: '⋃',
  0x56: '△',
  0x57: '⬜',
  0x61: '↦',
  0x68: 'ℏ',
  0x6c: 'ℓ',
  0xa1: 'ℝ',
}

const MT_CODE: Readonly<Record<number, string>> = {
  0xe083: '←',
  0xe091: '↔',
  0xe096: '→',
  0xe098: '→',
  0xe421: '⩾',
  0xe425: '⩽',
  0xe901: '⩲',
  0xe902: '⩱',
  0xe903: '⨦',
  0xe904: '⨤',
  0xe908: '⪜',
  0xe909: '⪛',
  0xe90a: '⋾',
  0xe90b: '⊇',
  0xe90c: '⊆',
  0xe90d: '⪕',
  0xe912: '⋷',
  0xe913: '⪚',
  0xe914: '⫈',
  0xe915: '⫇',
  0xe916: '⪾',
  0xe917: '⪽',
  0xe918: '⩦',
  0xe919: '⨫',
  0xe91a: '⨬',
  0xe91f: '⩬',
  0xe920: '⩏',
  0xe921: '⩎',
  0xe922: '⋚',
  0xe924: '⩪',
  0xe925: '⩫',
  0xe926: '⪋',
  0xe927: '⪌',
  0xe928: '≲',
  0xe929: '≳',
  0xe92a: '⫍',
  0xe92b: '⫎',
  0xe92c: '⩧',
  0xe92d: '⋛',
  0xe92e: '|',
  0xe92f: '‖',
  0xe930: '⫼',
  0xe932: '≲',
  0xe933: '≳',
  0xe936: '⪡',
  0xe937: '⪢',
  0xe93a: '⪷',
  0xe93b: '⪸',
  0xe940: '⪳',
  0xe941: '⪴',
  0xe948: '⫤',
  0xe949: '∾',
  0xe94a: '⊣',
  0xe94b: '⩵',
  0xe94c: '⩶',
  0xe94d: '⧴',
  0xe94e: '≡',
  0xe950: '⧏',
  0xe951: '⧐',
  0xe954: '⊃',
  0xe955: '⌣̅',
  0xe956: '⌢̅',
  0xe957: '⫊',
  0xe958: '⫉',
  0xe959: '⪐',
  0xe95a: '⪏',
  0xe95c: '⩓',
  0xe95d: '⩔',
  0xe95e: '⩣',
  0xe95f: '⩢',
  0xe962: '⩰',
  0xe966: '⩸',
  0xe96b: '⫷',
  0xe96c: '⫸',
  0xe96d: '≏',
  0xe96e: '⪤',
  0xe96f: '⩙',
  0xe970: '⫖',
  0xe971: '⫕',
  0xe972: '⫔',
  0xe973: '⫓',
  0xe979: '⫴',
  0xe97a: '⁞⁞',
  0xe97b: '⫨',
  0xe97c: '⫣',
  0xe97d: '⫥',
  0xe97e: '⫩',
  0xe97f: '⫤',
  0xe980: '⦡',
  0xe981: '⫽',
  0xe982: '⦜',
  0xe984: '⦶',
  0xe985: '⨸',
  0xe98a: '⫡',
  0xe98b: '⦞',
  0xe98c: '⦠',
  0xe98e: '⥾',
  0xe98f: '⦁',
  0xe990: '⧋',
  0xe992: '⨳',
  0xe994: '⧥',
  0xe995: '⩩',
  0xe996: '⫰',
  0xe998: '☾',
  0xe999: '☺',
  0xe9a5: '⬡',
  0xe9a6: '⨤',
  0xe9a7: '⨦',
  0xec07: '|',
  0xec08: '|',
  0xec09: '‖',
  0xec0a: '‖',
  0xed01: '⅁',
  0xed02: '𝚥',
  0xed03: 'ϝ',
  0xed04: 'ⅆ',
  0xed05: 'ⅇ',
  0xed06: 'ⅈ',
  0xed07: 'ⅉ',
  0xed08: 'ⅅ',
  0xef80: '∲',
  0xef81: '⨖',
  0xef82: '⨏',
}

const MT_SPACES = [
  '\u200b',
  '\u2009',
  '\u2005',
  '\u2004',
  '\u2003',
  '\u2003\u2003',
  '\u2009',
  '\u2002',
]

function fontScript(font: Typeface | undefined): ScriptFont | undefined {
  const family = `${font?.name ?? ''} ${font?.encoding ?? ''}`.toLowerCase()
  if (/fraktur|blackletter/.test(family)) return 'fraktur'
  if (/euclid\s*math\s*(two|2)|blackboard|double.struck/.test(family)) return 'double-struck'
  if (/euclid\s*math\s*(one|1)|calligraph|script/.test(family)) return 'script'
  if (/monospace|courier/.test(family)) return 'monospace'
  if (/sans.serif/.test(family)) return 'sans-serif'
}

function fontCharacter(
  position: number,
  typeface: number,
  font: Typeface | undefined,
): string | undefined {
  const family = `${font?.name ?? ''} ${font?.encoding ?? ''}`.toLowerCase()
  if (position >= 0xf000 && position <= 0xf0ff) position -= 0xf000
  if (/\b(?:m\s*t|euclid)\s*extra/.test(family) || (!font && typeface === 11))
    return MT_EXTRA[position]
  if (/symbol/.test(family) || (!font && typeface >= 4 && typeface <= 6)) {
    if (position >= 0x41 && position <= 0x5a) return SYMBOL_UPPER[position - 0x41]
    if (position >= 0x61 && position <= 0x7a) return SYMBOL_LOWER[position - 0x61]
    return SYMBOL_ANSI[position] ?? (position < 0x7f ? String.fromCharCode(position) : undefined)
  }
  if (/win/.test(family) && position >= 0x80 && position <= 0x9f) {
    return '€\u0000‚ƒ„…†‡ˆ‰Š‹Œ\u0000Ž\u0000\u0000‘’“”•–—˜™š›œ\u0000žŸ'[position - 0x80]
  }
  return position < 0xe000 || position > 0xf8ff ? String.fromCharCode(position) : undefined
}

function mappedChar(
  code: number | undefined,
  position: number | undefined,
  typeface: number,
  font: Typeface | undefined,
): string | undefined {
  requireValid(code !== undefined || position !== undefined)
  let char: string | undefined
  if (code !== undefined && (code < 0xe000 || code > 0xf8ff)) {
    char = String.fromCharCode(code)
  } else if (code !== undefined) {
    const page = code & 0xff00
    const index = code & 0xff
    if (page === 0xf100 && index < 52)
      char = String.fromCharCode(index < 26 ? 65 + index : 97 + index - 26)
    else if ((page === 0xeb00 || page === 0xef00) && index >= 1 && index <= MT_SPACES.length)
      char = MT_SPACES[index - 1]
    else if (page === 0xf000 && (font || typeface === 11 || (typeface >= 4 && typeface <= 6)))
      char = fontCharacter(code, typeface, font)
    else char = MT_CODE[code]
  }
  if (char === undefined && position !== undefined) char = fontCharacter(position, typeface, font)
  return char
}

const SCRIPT_LETTERS = '𝒜ℬ𝒞𝒟ℰℱ𝒢ℋℐ𝒥𝒦ℒℳ𝒩𝒪𝒫𝒬ℛ𝒮𝒯𝒰𝒱𝒲𝒳𝒴𝒵𝒶𝒷𝒸𝒹ℯ𝒻ℊ𝒽𝒾𝒿𝓀𝓁𝓂𝓃ℴ𝓅𝓆𝓇𝓈𝓉𝓊𝓋𝓌𝓍𝓎𝓏'
const FRAKTUR_LETTERS = '𝔄𝔅ℭ𝔇𝔈𝔉𝔊ℌℑ𝔍𝔎𝔏𝔐𝔑𝔒𝔓𝔔ℜ𝔖𝔗𝔘𝔙𝔚𝔛𝔜ℨ'
const DOUBLE_STRUCK_LETTERS = '𝔸𝔹ℂ𝔻𝔼𝔽𝔾ℍ𝕀𝕁𝕂𝕃𝕄ℕ𝕆ℙℚℝ𝕊𝕋𝕌𝕍𝕎𝕏𝕐ℤ'
const SCRIPT_INDEX = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'

function scriptLetter(char: string, script: ScriptFont | undefined): string {
  const alphabet =
    script === 'script'
      ? SCRIPT_LETTERS
      : script === 'fraktur'
        ? FRAKTUR_LETTERS
        : script === 'double-struck'
          ? DOUBLE_STRUCK_LETTERS
          : ''
  const mappedIndex = [...alphabet].indexOf(char)
  return mappedIndex >= 0 ? SCRIPT_INDEX[mappedIndex] : char
}

// Match xml-utils without importing its XML parser into the binary decoder.
function escapeText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function element(name: string, content: string): string {
  return content ? `<m:${name}>${content}</m:${name}>` : `<m:${name}/>`
}

function property(name: string, value: string): string {
  return `<m:${name} m:val="${escapeText(value).replace(/"/g, '&quot;')}"/>`
}

function run(text: string, style?: Style, script?: ScriptFont): string {
  const props = (script ? property('scr', script) : '') + (style ? property('sty', style) : '')
  const space = /^\s|\s$/u.test(text) ? ' xml:space="preserve"' : ''
  return element(
    'r',
    (props ? element('rPr', props) : '') + `<m:t${space}>${escapeText(text)}</m:t>`,
  )
}

type CharNode = {
  kind: 'char'
  text?: string
  style?: Style
  font?: ScriptFont
  joinText: boolean
  embellishments: number[]
}
type SlotNode = { kind: 'line'; xml: string } | { kind: 'pile'; xml: string }
type ScriptNode = { kind: 'script'; sub?: string; sup?: string; pre: boolean }
type Node =
  | CharNode
  | SlotNode
  | ScriptNode
  | { kind: 'math'; xml: string }
  | { kind: 'embell'; code: number }

function structure(name: string, props: string, content: string): string {
  return element(name, (props ? element(`${name}Pr`, props) : '') + content)
}

function scriptXml(base: string, { pre, sub, sup }: ScriptNode): string {
  const inner = element('e', base)
  const lower = element('sub', sub ?? '')
  const upper = element('sup', sup ?? '')
  if (pre) return element('sPre', lower + upper + inner)
  if (sub === undefined) return element('sSup', inner + upper)
  if (sup === undefined) return element('sSub', inner + lower)
  return element('sSubSup', inner + lower + upper)
}

function accent(base: string, char: string): string {
  return structure('acc', property('chr', char), element('e', base))
}

function group(base: string, char: string, above: boolean): string {
  const pos = above ? 'top' : 'bot'
  return structure(
    'groupChr',
    property('chr', char) + property('pos', pos) + property('vertJc', above ? 'bot' : 'top'),
    element('e', base),
  )
}

function bar(base: string, above: boolean, double = false): string {
  if (double) return group(base, '‗', above)
  return structure('bar', property('pos', above ? 'top' : 'bot'), element('e', base))
}

function limit(base: string, label: string, above: boolean): string {
  return element(above ? 'limUpp' : 'limLow', element('e', base) + element('lim', label))
}

function box(base: string, props: string): string {
  return structure('borderBox', props, element('e', base))
}

const OVER_ACCENTS: Readonly<Record<number, string>> = {
  2: '\u0307',
  3: '\u0308',
  4: '\u20db',
  8: '\u0303',
  9: '\u0302',
  11: '\u20d7',
  12: '\u20d6',
  13: '\u20e1',
  14: '⇀',
  15: '↼',
  19: '⌢',
  20: '⌣',
  24: '\u20dc',
}
const UNDER_ACCENTS: Readonly<Record<number, string>> = {
  25: '˙',
  26: '¨',
  27: '\u20db',
  28: '\u20dc',
  30: '~',
  31: '⌢',
  32: '⌣',
  33: '→',
  34: '←',
  35: '↔',
  36: '⇀',
  37: '↼',
}
const PRIMES: Readonly<Record<number, string>> = { 5: '′', 6: '″', 7: '‵', 18: '‴' }
const STRIKES: Readonly<Record<number, string[]>> = {
  10: ['BLTR'],
  16: ['H'],
  21: ['BLTR', 'TLBR'],
  22: ['BLTR'],
  23: ['TLBR'],
}

function embellish(base: string, code: number): string {
  const prime = PRIMES[code]
  if (prime) return scriptXml(base, { kind: 'script', pre: code === 7, sup: run(prime) })
  if (code === 17 || code === 29) return bar(base, code === 17)
  if (OVER_ACCENTS[code]) return accent(base, OVER_ACCENTS[code])
  if (UNDER_ACCENTS[code]) return group(base, UNDER_ACCENTS[code], false)
  if (STRIKES[code])
    return box(
      base,
      HIDDEN_SIDES + STRIKES[code].map((way) => property(`strike${way}`, '1')).join(''),
    )
  throw new InvalidEquation()
}

type Atom = { xml: string; script?: ScriptNode; height: number }

function lineXml(nodes: Node[]): string {
  const atoms: Atom[] = []
  let pending: ScriptNode | undefined
  let text = ''
  const rendered = ({ xml, script }: Atom) => (script ? scriptXml(xml, script) : xml)
  const flushText = () => {
    if (text) atoms.push({ xml: run(text, 'p'), height: 0 })
    text = ''
  }
  const compatible = (first: ScriptNode, second: ScriptNode) =>
    (first.sub === undefined || second.sub === undefined) &&
    (first.sup === undefined || second.sup === undefined)
  const combine = (a: ScriptNode, b: ScriptNode): ScriptNode => ({
    ...a,
    sub: a.sub ?? b.sub,
    sup: a.sup ?? b.sup,
  })

  for (const node of nodes) {
    if (node.kind === 'script') {
      flushText()
      if (node.pre) {
        requireValid(!pending || compatible(pending, node))
        pending = pending ? combine(pending, node) : node
      } else {
        const base = atoms.pop()
        requireValid(base && !pending && base.height < MAX_DEPTH)
        const merge = base.script && !base.script.pre && compatible(base.script, node)
        if (merge) atoms.push({ ...base, script: combine(base.script!, node) })
        else atoms.push({ xml: rendered(base), script: node, height: base.height + 1 })
      }
      continue
    }
    requireValid(node.kind !== 'embell')
    let xml: string
    if (node.kind === 'char') {
      requireValid(node.text)
      if (
        !pending &&
        !node.embellishments.length &&
        node.joinText &&
        node.style === 'p' &&
        !node.font
      ) {
        text += node.text
        continue
      }
      xml = node.embellishments.reduce(embellish, run(node.text, node.style, node.font))
    } else {
      xml = node.xml
    }
    flushText()
    if (xml) {
      atoms.push({ xml, script: pending, height: pending ? 1 : 0 })
      pending = undefined
    }
  }
  requireValid(!pending)
  flushText()
  return atoms.map(rendered).join('')
}

const FENCES = ['⟨⟩', '()', '{}', '[]', '||', '‖‖', '⌊⌋', '⌈⌉', '⟦⟧']
const LEGACY_INTERVALS = ['[[', ']]', '][', '[)', '(]']
const BIG_OPERATORS = ['∑', '∏', '∐', '⋃', '⋂']
const INTEGRAL_SIGNS = ['∫', '∬', '∭', '∮', '∯', '∰', '∲', '∳']
const HIDDEN_SIDES = ['hideTop', 'hideBot', 'hideLeft', 'hideRight']
  .map((name) => property(name, '1'))
  .join('')

function delimit(slots: string[], left: string, right: string, separator = ''): string {
  return structure(
    'd',
    property('begChr', left) + property('sepChr', separator) + property('endChr', right),
    slots.map((slot) => element('e', slot)).join(''),
  )
}

function nary(
  slots: string[],
  operator: string | undefined,
  lower: boolean,
  upper: boolean,
  sumStyle: boolean,
): string {
  requireValid(operator && slots.length >= 1 && slots.length <= 3)
  const [body, first = '', second = ''] = slots
  // Exporters can omit inactive slots or write explicit null lines in their places.
  const sub = lower ? first : ''
  const sup = upper ? (slots.length === 2 && !lower ? first : second) : ''
  requireValid((lower || (upper && slots.length === 2) || !first) && (upper || !second))
  const props =
    property('chr', operator) +
    property('limLoc', sumStyle ? 'undOvr' : 'subSup') +
    property('subHide', lower ? '0' : '1') +
    property('supHide', upper ? '0' : '1')
  return structure('nary', props, element('sub', sub) + element('sup', sup) + element('e', body))
}

function fraction(slots: string[], type?: 'lin' | 'skw'): string {
  requireValid(slots.length === 2)
  return structure(
    'f',
    type ? property('type', type) : '',
    element('num', slots[0]) + element('den', slots[1]),
  )
}

function radical(slots: string[], degree: boolean): string {
  requireValid(slots.length >= 1 && slots.length <= 2 && (degree || !slots[1]))
  return structure(
    'rad',
    property('degHide', degree ? '0' : '1'),
    element('deg', slots[1] ?? '') + element('e', slots[0]),
  )
}

function scriptTemplate(slots: string[], sub: boolean, sup: boolean, pre: boolean): ScriptNode {
  requireValid(slots.length >= 1 && slots.length <= 2 && (slots.length === 2 || !sub || !sup))
  requireValid((sub || slots.length === 1 || !slots[0]) && (sup || !slots[1]))
  return {
    kind: 'script',
    sub: sub ? slots[0] : undefined,
    sup: sup ? slots[slots.length - 1] : undefined,
    pre,
  }
}

function limits(slots: string[], lower: boolean, upper: boolean): string {
  requireValid(slots.length >= 1 && slots.length <= 3)
  const [base, first = '', second = ''] = slots
  requireValid((lower || (upper && slots.length === 2) || !first) && (upper || !second))
  let xml = base
  if (lower) xml = limit(xml, first, false)
  if (upper) xml = limit(xml, slots.length === 2 && !lower ? first : second, true)
  return xml
}

function arrow(slots: string[], glyph: string, top: boolean, bottom: boolean): string {
  requireValid(slots.length <= 2)
  let xml = run(glyph)
  if (bottom) xml = limit(xml, slots[top || slots.length === 2 ? 1 : 0] ?? '', false)
  if (top) xml = limit(xml, slots[0] ?? '', true)
  return xml
}

function vector(variation: number): string {
  const direction = variation & 3
  if (variation & 8) return direction === 1 ? '↼' : direction === 3 ? '⇌' : '⇀'
  return direction === 1 ? '←' : direction === 3 ? '↔' : '→'
}

function dirac(
  slots: string[],
  chars: (string | undefined)[],
  left: boolean,
  right: boolean,
): string {
  requireValid((left || right) && slots.length >= 1 && slots.length <= 2)
  const first = left ? (chars[0] ?? '⟨') : ''
  const sep = chars[left ? 1 : 0] ?? '|'
  const last = right ? (chars[left ? 2 : 1] ?? '⟩') : ''
  if (left && right) {
    requireValid(slots.length === 2)
    return delimit(slots, first, last, sep)
  }
  return delimit([slots[right ? slots.length - 1 : 0]], first || sep, last || sep)
}

function v5Template(
  selector: number,
  variation: number,
  slots: string[],
  chars: (string | undefined)[],
): string | ScriptNode {
  const [main] = slots
  const glyph = chars[0]
  if (selector <= 9) {
    requireValid(slots.length === 1)
    if (selector === 9) {
      const ends = '()[]'
      return delimit(slots, glyph ?? ends[variation & 3], chars[1] ?? ends[(variation >> 4) & 3])
    }
    const left = (variation & 1) !== 0
    const right = (variation & 2) !== 0
    return delimit(
      slots,
      left ? (glyph ?? FENCES[selector][0]) : '',
      right ? (chars[left ? 1 : 0] ?? FENCES[selector][1]) : '',
    )
  }
  if (selector === 10) return radical(slots, (variation & 1) !== 0)
  if (selector === 11)
    return fraction(slots, variation & 2 ? (variation & 4 ? 'lin' : 'skw') : undefined)
  if (selector === 12 || selector === 13) {
    requireValid(slots.length === 1)
    return bar(main, selector === 13, (variation & 1) !== 0)
  }
  if (selector === 14) {
    const direction = (variation >> 4) & 3
    const type = variation & 3
    const sign =
      type === 1
        ? direction === 1
          ? '⇆'
          : '⇄'
        : type === 2
          ? direction === 1
            ? '⇋'
            : '⇌'
          : direction === 1
            ? '←'
            : direction === 3
              ? '↔'
              : '→'
    return arrow(slots, glyph ?? sign, (variation & 4) !== 0, (variation & 8) !== 0)
  }
  if (selector >= 15 && selector <= 23) {
    // MathType writers use the upper nibble for limits. Older documentation
    // identifies two low bits which are also integral sign variants.
    const limitBits = variation & 0x30 || (selector !== 15 ? (variation & 3) << 4 : 0)
    const lower = (limitBits & 0x10) !== 0
    const upper = (limitBits & 0x20) !== 0
    if (selector === 23) return limits(slots, lower, upper)
    const integralCount = Math.max(1, variation & 3)
    const loop = (variation >> 2) & 3
    const integralSign = INTEGRAL_SIGNS[loop < 2 ? loop * 3 + integralCount - 1 : loop + 4]
    // A double or triple integral can contain several single-integral glyphs.
    const operator = selector === 15 ? integralSign : (BIG_OPERATORS[selector - 16] ?? glyph)
    return nary(slots, operator, lower, upper, (variation & 0x40) !== 0)
  }
  if (selector === 24 || selector === 25) {
    requireValid(slots.length >= 1 && slots.length <= 2)
    const above = (variation & 1) !== 0
    const sign = selector === 24 ? (above ? '⏞' : '⏟') : above ? '⎴' : '⎵'
    const grouped = group(main, sign, above)
    return slots[1] ? limit(grouped, slots[1], above) : grouped
  }
  if (selector >= 27 && selector <= 29)
    return scriptTemplate(slots, selector !== 28, selector !== 27, (variation & 1) !== 0)
  if (selector === 30) return dirac(slots, chars, (variation & 1) !== 0, (variation & 2) !== 0)
  requireValid(slots.length === 1)
  if (selector === 31) {
    if (variation & 4) return group(main, vector(variation), false)
    return accent(main, (variation & 0xf) === 2 ? '\u20d7' : vector(variation))
  }
  if (selector >= 32 && selector <= 34)
    return accent(main, ['\u0303', '\u0302', '⌢'][selector - 32])
  if (selector === 37) {
    const props = Object.entries({ Top: 8, Bot: 16, Left: 2, Right: 4 })
      .filter(([, flag]) => !(variation & flag))
      .map(([side]) => property(`hide${side}`, '1'))
      .join('')
    return box(main, props)
  }
  throw new InvalidEquation()
}

function v3Template(
  selector: number,
  variation: number,
  slots: string[],
  chars: (string | undefined)[],
): string | ScriptNode {
  const [main] = slots
  const glyph = chars[0]
  if (selector <= 12) {
    requireValid(slots.length === 1)
    const ends = selector <= 7 ? FENCES[selector] : LEGACY_INTERVALS[selector - 8]
    const left = selector >= 6 || variation !== 2
    const right = selector >= 6 || variation !== 1
    return delimit(
      slots,
      left ? (glyph ?? ends[0]) : '',
      right ? (chars[left ? 1 : 0] ?? ends[1]) : '',
    )
  }
  if (selector === 13) return radical(slots, variation === 1)
  if (selector === 14 || selector === 41)
    return fraction(slots, selector === 41 ? (variation === 1 ? 'lin' : 'skw') : undefined)
  if (selector === 15 || selector === 44)
    return scriptTemplate(slots, variation !== 0, variation !== 1, selector === 44)
  if (selector === 16 || selector === 17) {
    requireValid(slots.length === 1)
    return bar(main, selector === 17, variation === 1)
  }
  if (selector >= 18 && selector <= 20)
    return arrow(slots, glyph ?? ['←', '→', '↔'][selector - 18], variation === 0, variation === 1)
  if (selector >= 21 && selector <= 26) {
    const shifted = selector >= 24
    const offset = (selector - 21) % 3
    const loop = shifted
      ? offset === 0
        ? variation === 2
        : variation === 0
      : offset === 0
        ? variation >= 3
        : variation >= 2
    const lower = shifted || (offset === 0 ? [1, 2, 4].includes(variation) : (variation & 1) !== 0)
    const upper = offset === 0 && variation === (shifted ? 0 : 2)
    return nary(slots, INTEGRAL_SIGNS[offset + (loop ? 3 : 0)], lower, upper, shifted)
  }
  if (selector === 27 || selector === 28)
    return v5Template(24, selector === 27 ? 1 : 0, slots, chars)
  if (selector >= 29 && selector <= 38)
    return nary(
      slots,
      BIG_OPERATORS[Math.floor((selector - 29) / 2)],
      variation !== 2,
      variation === 1,
      (selector & 1) !== 0,
    )
  if (selector === 39) return limits(slots, variation !== 0, variation !== 1)
  if (selector === 42 || selector === 43)
    return nary(slots, glyph, variation !== 0, variation !== 1, selector === 43)
  if (selector === 45) return dirac(slots, chars, variation !== 2, variation !== 1)
  requireValid(slots.length === 1)
  if (selector === 46 || selector === 47)
    return group(main, ['←', '→', '↔'][variation] ?? '→', selector === 47)
  if (selector === 48) return accent(main, '⌢')
  throw new InvalidEquation()
}

class MtefReader {
  private offset = 0
  private records = 0
  private readonly legacy: boolean
  private readonly encodings = ['', 'MTCode', 'Unknown', 'Symbol', 'MTExtra']
  private readonly fonts: { name: string; encoding: string }[] = []
  private readonly styles: Typeface[] = []
  private readonly preferences: (Typeface | undefined)[] = []

  constructor(
    private readonly bytes: Uint8Array,
    private readonly alternateRecords: boolean,
  ) {
    const version = this.byte()
    requireValid(version === 3 || version === 5)
    this.legacy = version === 3
    requireValid(this.byte() <= 1 && this.byte() <= 1)
    this.skip(2)
    if (!this.legacy) {
      this.cString()
      requireValid(this.byte() <= 1)
    }
  }

  private byte(): number {
    requireValid(this.offset < this.bytes.length)
    return this.bytes[this.offset++]
  }

  private word(): number {
    return this.byte() | (this.byte() << 8)
  }

  private unsigned(): number {
    const first = this.byte()
    return first === 255 ? this.word() : first
  }

  private skip(count: number): void {
    requireValid(count <= this.bytes.length - this.offset)
    this.offset += count
  }

  private cString(): string {
    let out = ''
    for (let i = 0; i < 1024; i++) {
      const code = this.byte()
      if (code === 0) return out
      out += String.fromCharCode(code)
    }
    throw new InvalidEquation()
  }

  private bump(): void {
    requireValid(++this.records <= MAX_RECORDS)
  }

  private tag(): number {
    this.bump()
    return this.byte()
  }

  private nudge(options: number): void {
    if (options & 8) {
      const x = this.byte()
      const y = this.byte()
      if (x === 128 && y === 128) this.skip(4)
    }
  }

  private ruler(): void {
    this.skip(this.byte() * 3)
  }

  private rulerIfPresent(options: number): void {
    if (options & 2) {
      requireValid(this.tag() === 7)
      this.ruler()
    }
  }

  private dimensions(): void {
    let remaining = this.byte()
    while (remaining > 0) {
      const pair = this.byte()
      if (pair >> 4 === 15) remaining--
      if (remaining > 0 && (pair & 15) === 15) remaining--
    }
  }

  private fontStyle(id: number, flags: number): Typeface {
    requireValid(id > 0 && id <= this.fonts.length && flags <= 3)
    return { ...this.fonts[id - 1], flags }
  }

  private equationPreferences(): void {
    requireValid(this.byte() === 0)
    this.dimensions()
    this.dimensions()
    const count = this.byte()
    for (let i = 1; i <= count; i++) {
      const font = this.unsigned()
      this.preferences[i] = font ? this.fontStyle(font, this.byte()) : undefined
    }
  }

  private colorDefinition(): void {
    const options = this.byte()
    requireValid(options <= 7)
    const count = options & 1 ? 4 : 3
    for (let i = 0; i < count; i++) requireValid(this.word() <= 1000)
    if (options & 4) this.cString()
  }

  private line(options: number, depth: number): SlotNode {
    if (options & 4) this.skip(2)
    this.rulerIfPresent(options)
    return { kind: 'line', xml: options & 1 ? '' : lineXml(this.list(depth + 1)) }
  }

  private char(options: number, depth: number): CharNode {
    const typeface = this.legacy ? this.byte() - 128 : this.signedTypeface()
    const font = typeface < 0 ? this.styles[-typeface - 1] : this.preferences[typeface]
    const code = this.legacy || !(options & 0x20) ? this.word() : undefined
    let position: number | undefined
    if (!this.legacy) {
      requireValid((options & 0x14) !== 0x14)
      if (options & 4) position = this.byte()
      if (options & 0x10) position = this.word()
    }
    let text: string | undefined
    if (
      this.legacy &&
      code !== undefined &&
      code <= 0xff &&
      (typeface === 4 ||
        typeface === 5 ||
        typeface === 11 ||
        /symbol|extra/i.test(font?.name ?? ''))
    ) {
      text = fontCharacter(code, typeface, font) ?? mappedChar(code, undefined, typeface, font)
    } else {
      text = mappedChar(code, position, typeface, font)
    }
    requireValid(text === undefined || !hasInvalidXmlChar(text))
    const defaults = typeface === 3 || typeface === 4 ? 2 : typeface === 7 ? 1 : 0
    const flags = font?.flags ?? defaults
    let script = fontScript(font)
    if (code !== undefined && code >= 0xf100 && code <= 0xf133) script ??= 'script'
    const identifier = text !== undefined && /^[\p{L}\p{M}\s]+$/u.test(text)
    const upright = identifier && typeface !== 6 && typeface !== 11
    const style =
      flags === 0 && !script && !upright ? undefined : (['p', 'b', 'i', 'bi'] as const)[flags]
    const embellishments: number[] = []
    if (options & (this.legacy ? 2 : 1)) {
      for (const child of this.list(depth + 1)) {
        requireValid(child.kind === 'embell' && embellishments.length < MAX_DEPTH)
        embellishments.push(child.code)
      }
    }
    return {
      kind: 'char',
      text: text === undefined ? undefined : scriptLetter(text, script),
      style,
      font: script,
      joinText: typeface === 1 || typeface === 2 || typeface === 12,
      embellishments,
    }
  }

  private signedTypeface(): number {
    const first = this.byte()
    return first === 255 ? this.word() - 32768 : first - 128
  }

  private template(depth: number): Node {
    const selector = this.byte()
    let variation = this.byte()
    if (!this.legacy && variation & 0x80) variation = (variation & 0x7f) | (this.byte() << 8)
    this.byte()
    const slots: string[] = []
    const chars: (string | undefined)[] = []
    for (const child of this.list(depth + 1)) {
      if (child.kind === 'char') {
        requireValid(
          !child.embellishments.length &&
            (child.text === undefined || [...child.text].length === 1),
        )
        chars.push(child.text)
      } else {
        requireValid(child.kind === 'line' || child.kind === 'pile')
        slots.push(child.xml)
      }
    }
    const parsed = this.legacy
      ? v3Template(selector, variation, slots, chars)
      : v5Template(selector, variation, slots, chars)
    return typeof parsed === 'string' ? { kind: 'math', xml: parsed } : parsed
  }

  private pile(options: number, depth: number): SlotNode {
    this.byte()
    const vertical = this.byte()
    this.rulerIfPresent(options)
    const lines = this.list(depth + 1)
    requireValid(lines.length > 0 && lines.every((line) => line.kind === 'line'))
    const align = ['top', 'center', 'bot'][vertical] ?? 'center'
    return {
      kind: 'pile',
      xml: structure(
        'eqArr',
        property('baseJc', align),
        lines.map((line) => element('e', line.xml)).join(''),
      ),
    }
  }

  private matrix(depth: number): Node {
    this.byte()
    const horizontal = this.byte()
    this.byte()
    const rows = this.byte()
    const columns = this.byte()
    requireValid(rows > 0 && columns > 0 && rows * columns <= MAX_RECORDS)
    this.skip(Math.ceil((rows + 1) / 4) + Math.ceil((columns + 1) / 4))
    const cells = this.list(depth + 1)
    requireValid(
      cells.length === rows * columns &&
        cells.every((cell) => cell.kind === 'line' || cell.kind === 'pile'),
    )
    const align = horizontal === 1 ? 'left' : horizontal === 3 ? 'right' : 'center'
    const props = element(
      'mcs',
      element('mc', element('mcPr', property('count', String(columns)) + property('mcJc', align))),
    )
    let xml = ''
    for (let row = 0; row < rows; row++) {
      xml += element(
        'mr',
        cells
          .slice(row * columns, (row + 1) * columns)
          .map((cell) => element('e', cell.xml))
          .join(''),
      )
    }
    return { kind: 'math', xml: structure('m', props, xml) }
  }

  private list(depth: number): Node[] {
    requireValid(depth <= MAX_DEPTH)
    const nodes: Node[] = []
    for (;;) {
      const tag = this.tag()
      let type = this.legacy ? tag & 15 : tag
      if (type === 0) return nodes
      if (type >= 1 && type <= 6) {
        const options = this.legacy ? tag >> 4 : this.byte()
        this.nudge(options)
        switch (type) {
          case 1:
            nodes.push(this.line(options, depth))
            break
          case 2:
            nodes.push(this.char(options, depth))
            break
          case 3:
            nodes.push(this.template(depth))
            break
          case 4:
            nodes.push(this.pile(options, depth))
            break
          case 5:
            nodes.push(this.matrix(depth))
            break
          case 6:
            nodes.push({ kind: 'embell', code: this.byte() })
            if (this.legacy && !this.alternateRecords) this.skip(1)
            break
        }
        continue
      }
      if (!this.legacy && this.alternateRecords) {
        if (type === 15 || type === 16) continue
        if (type === 18 || type === 19) type -= 3
        else if (type === 20 || type === 21) type -= 2
      }
      switch (type) {
        case 7:
          this.ruler()
          break
        case 8: {
          if (this.legacy) {
            const id = this.byte()
            const flags = this.byte()
            requireValid(id > 0 && flags <= 3)
            this.styles[id - 1] = {
              name: this.cString(),
              encoding: '',
              flags: ((flags & 1) << 1) | (flags >> 1),
            }
          } else {
            this.styles.push(this.fontStyle(this.unsigned(), this.byte()))
          }
          break
        }
        case 9: {
          const size = this.byte()
          this.skip(size === 100 ? 3 : size === 101 ? 2 : 1)
          break
        }
        case 10:
        case 11:
        case 12:
        case 13:
        case 14:
          break
        default: {
          requireValid(!this.legacy)
          switch (type) {
            case 15:
              this.unsigned()
              break
            case 16:
              this.colorDefinition()
              break
            case 17: {
              const encoding = this.encodings[this.unsigned()]
              requireValid(encoding)
              this.fonts.push({ encoding, name: this.cString() })
              break
            }
            case 18:
              this.equationPreferences()
              break
            case 19:
              this.encodings.push(this.cString())
              break
            default:
              requireValid(type >= 100)
              this.skip(this.unsigned())
          }
        }
      }
    }
  }

  toOmml(): string {
    const nodes = this.list(0)
    requireValid(this.bytes.subarray(this.offset).every((byte) => byte === 0))
    const xml = lineXml(nodes)
    requireValid(xml)
    return element('oMath', xml)
  }
}

export function mtefToOmml(mtefBytes: Uint8Array): string | null {
  if (mtefBytes.length > MAX_MTEF_BYTES || (mtefBytes[0] !== 3 && mtefBytes[0] !== 5)) return null
  // Equation Editor pads v3 embellishments by one byte. Also accept the
  // documented unpadded form, and both sets of v5 definition record IDs.
  for (const alternateRecords of [false, true]) {
    try {
      return new MtefReader(mtefBytes, alternateRecords).toOmml()
    } catch (error) {
      if (!(error instanceof InvalidEquation)) throw error
    }
  }
  return null
}

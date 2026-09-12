const CP1252: Record<number, string> = {
  0x80: '€',
  0x82: '‚',
  0x83: 'ƒ',
  0x84: '„',
  0x85: '…',
  0x86: '†',
  0x87: '‡',
  0x88: 'ˆ',
  0x89: '‰',
  0x8a: 'Š',
  0x8b: '‹',
  0x8c: 'Œ',
  0x8e: 'Ž',
  0x91: '‘',
  0x92: '’',
  0x93: '“',
  0x94: '”',
  0x95: '•',
  0x96: '–',
  0x97: '—',
  0x98: '˜',
  0x99: '™',
  0x9a: 'š',
  0x9b: '›',
  0x9c: 'œ',
  0x9e: 'ž',
  0x9f: 'Ÿ',
}

const DESTINATIONS = new Set([
  'colortbl',
  'datastore',
  'filetbl',
  'fonttbl',
  'footer',
  'footerf',
  'footerl',
  'footerr',
  'header',
  'headerf',
  'headerl',
  'headerr',
  'info',
  'listoverridetable',
  'listtable',
  'pict',
  'revtbl',
  'stylesheet',
  'themedata',
])

const SYMBOLS: Record<string, string> = {
  bullet: '•',
  emdash: '—',
  endash: '–',
  lquote: '‘',
  rquote: '’',
  ldblquote: '“',
  rdblquote: '”',
  line: '\n',
  par: '\n',
  tab: '\t',
}

interface RtfState {
  skip: boolean
  uc: number
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  vertAlign?: 'superscript' | 'subscript'
  caps?: 'all' | 'small'
  sizeHalfPoints?: number
  indentLeft?: number
  indentRight?: number
  indentFirstLine?: number
  spaceBefore?: number
  spaceAfter?: number
  align?: 'left' | 'center' | 'right' | 'justify'
}

export interface ZoteroRtfRun {
  text: string
  bold?: true
  italic?: true
  underline?: true
  strike?: true
  vertAlign?: 'superscript' | 'subscript'
  caps?: 'all' | 'small'
  sizeHalfPoints?: number
}

export interface ZoteroRtfParagraph {
  runs: ZoteroRtfRun[]
  indentLeft?: number
  indentRight?: number
  indentFirstLine?: number
  spaceBefore?: number
  spaceAfter?: number
  align?: 'left' | 'center' | 'right' | 'justify'
}

export interface ZoteroRtfDocument {
  paragraphs: ZoteroRtfParagraph[]
}

function plainDocument(input: string): ZoteroRtfDocument {
  return {
    paragraphs: input.split(/\r?\n/).map((text) => ({ runs: text ? [{ text }] : [] })),
  }
}

/** Parse the formatting subset emitted by Zotero's LibreOffice integration. */
export function parseZoteroRtf(input: string): ZoteroRtfDocument {
  if (!/^\{\\rtf/i.test(input.trim())) return plainDocument(input)
  const initial: RtfState = {
    skip: false,
    uc: 1,
    bold: false,
    italic: false,
    underline: false,
    strike: false,
  }
  const states: RtfState[] = [initial]
  let state = states[0]
  const paragraphs: ZoteroRtfParagraph[] = []
  let paragraph: ZoteroRtfParagraph = { runs: [] }
  let skipFallback = 0

  const syncParagraphFormat = () => {
    for (const key of [
      'indentLeft',
      'indentRight',
      'indentFirstLine',
      'spaceBefore',
      'spaceAfter',
      'align',
    ] as const) {
      const value = state[key]
      if (value === undefined) delete paragraph[key]
      else Object.assign(paragraph, { [key]: value })
    }
  }
  const resetParagraph = () => {
    paragraphs.push(paragraph)
    paragraph = { runs: [] }
    syncParagraphFormat()
  }
  const append = (text: string) => {
    if (!text) return
    const run: ZoteroRtfRun = {
      text,
      ...(state.bold ? { bold: true } : {}),
      ...(state.italic ? { italic: true } : {}),
      ...(state.underline ? { underline: true } : {}),
      ...(state.strike ? { strike: true } : {}),
      ...(state.vertAlign ? { vertAlign: state.vertAlign } : {}),
      ...(state.caps ? { caps: state.caps } : {}),
      ...(state.sizeHalfPoints ? { sizeHalfPoints: state.sizeHalfPoints } : {}),
    }
    const previous = paragraph.runs[paragraph.runs.length - 1]
    const style = ({ text: _text, ...attrs }: ZoteroRtfRun) => JSON.stringify(attrs)
    if (previous && style(previous) === style(run)) previous.text += text
    else paragraph.runs.push(run)
  }
  const appendDecoded = (text: string) => {
    if (skipFallback > 0) skipFallback--
    else append(text)
  }

  for (let i = 0; i < input.length; i++) {
    const char = input[i]
    if (char === '{') {
      state = { ...state }
      states.push(state)
      continue
    }
    if (char === '}') {
      if (states.length > 1) states.pop()
      state = states[states.length - 1]
      continue
    }
    if (char !== '\\') {
      if (!state.skip && char !== '\r' && char !== '\n') {
        appendDecoded(char)
      }
      continue
    }

    const next = input[++i]
    if (next === undefined) break
    if (next === '\\' || next === '{' || next === '}') {
      if (!state.skip) appendDecoded(next)
      continue
    }
    if (next === '\r' || next === '\n') {
      // Zotero joins citeproc bibliography entries with a backslash plus CRLF.
      if (!state.skip) resetParagraph()
      if (next === '\r' && input[i + 1] === '\n') i++
      continue
    }
    if (next === '*') {
      state.skip = true
      continue
    }
    if (next === "'") {
      const hex = input.slice(i + 1, i + 3)
      if (/^[0-9a-f]{2}$/i.test(hex)) {
        i += 2
        if (!state.skip) {
          const value = parseInt(hex, 16)
          appendDecoded(CP1252[value] ?? String.fromCharCode(value))
        }
      }
      continue
    }
    if (!/[A-Za-z]/.test(next)) {
      if (!state.skip && next === '~') append('\u00a0')
      else if (!state.skip && next === '_') append('\u2011')
      continue
    }

    let word = next
    while (i + 1 < input.length && /[A-Za-z]/.test(input[i + 1])) word += input[++i]
    let sign = 1
    if (input[i + 1] === '-') {
      sign = -1
      i++
    }
    let digits = ''
    while (i + 1 < input.length && /\d/.test(input[i + 1])) digits += input[++i]
    const parameter = digits ? sign * Number(digits) : null
    if (input[i + 1] === ' ') i++

    const lower = word.toLowerCase()
    if (DESTINATIONS.has(lower)) {
      state.skip = true
      continue
    }
    if (lower === 'uc' && parameter !== null) {
      state.uc = Math.max(0, parameter)
      continue
    }
    if (lower === 'u' && parameter !== null && !state.skip) {
      append(String.fromCharCode(parameter < 0 ? parameter + 65536 : parameter))
      skipFallback = state.uc
      continue
    }
    if (state.skip) continue
    if (lower === 'par' || lower === 'line') {
      resetParagraph()
      continue
    }
    if (lower === 'tab') {
      append('\t')
      continue
    }
    if (SYMBOLS[lower] !== undefined) {
      append(SYMBOLS[lower])
      continue
    }
    if (lower === 'b') state.bold = parameter !== 0
    else if (lower === 'i') state.italic = parameter !== 0
    else if (lower === 'ul') state.underline = parameter !== 0
    else if (lower === 'ulnone') state.underline = false
    else if (lower === 'strike') state.strike = parameter !== 0
    else if (lower === 'super') state.vertAlign = 'superscript'
    else if (lower === 'sub') state.vertAlign = 'subscript'
    else if (lower === 'nosupersub') state.vertAlign = undefined
    else if (lower === 'caps') state.caps = parameter === 0 ? undefined : 'all'
    else if (lower === 'scaps') state.caps = parameter === 0 ? undefined : 'small'
    else if (lower === 'fs' && parameter !== null) state.sizeHalfPoints = parameter
    else if (lower === 'plain') {
      state.bold = false
      state.italic = false
      state.underline = false
      state.strike = false
      state.vertAlign = undefined
      state.caps = undefined
      state.sizeHalfPoints = undefined
    } else if (lower === 'pard') {
      state.indentLeft = undefined
      state.indentRight = undefined
      state.indentFirstLine = undefined
      state.spaceBefore = undefined
      state.spaceAfter = undefined
      state.align = undefined
      syncParagraphFormat()
    } else if (lower === 'li' && parameter !== null) {
      state.indentLeft = parameter
      syncParagraphFormat()
    } else if (lower === 'ri' && parameter !== null) {
      state.indentRight = parameter
      syncParagraphFormat()
    } else if (lower === 'fi' && parameter !== null) {
      state.indentFirstLine = parameter
      syncParagraphFormat()
    } else if (lower === 'sb' && parameter !== null) {
      state.spaceBefore = parameter
      syncParagraphFormat()
    } else if (lower === 'sa' && parameter !== null) {
      state.spaceAfter = parameter
      syncParagraphFormat()
    } else if (lower === 'ql' || lower === 'qc' || lower === 'qr' || lower === 'qj') {
      state.align = { ql: 'left', qc: 'center', qr: 'right', qj: 'justify' }[lower] as
        'left' | 'center' | 'right' | 'justify'
      syncParagraphFormat()
    }
  }

  paragraphs.push(paragraph)
  while (paragraphs.length > 1 && paragraphs[0].runs.length === 0) paragraphs.shift()
  while (paragraphs.length > 1 && paragraphs[paragraphs.length - 1].runs.length === 0) {
    paragraphs.pop()
  }
  return { paragraphs }
}

/** Convert Zotero's LibreOffice RTF output to editable Unicode text. */
export function zoteroRtfToText(input: string): string {
  if (!/^\{\\rtf/i.test(input.trim())) return input
  return parseZoteroRtf(input)
    .paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
}

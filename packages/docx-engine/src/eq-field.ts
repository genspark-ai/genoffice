// Legacy EQ field instructions (`EQ \f(a,b)`, `EQ \a \al (x,y)` ...) laid out
// as OMML so the math renderer draws them; Word computes these fields live and
// stores no cached result, so without this the field shows nothing.
import { decodeEntities } from './parse-xml-text'
import { escapeXmlAttr, escapeXmlText } from './xml-utils'

export interface EqField {
  omml: string
  /** linear text of the layout (previews, word count) */
  text: string
}

interface Piece {
  omml: string
  text: string
}

interface Parser {
  src: string
  i: number
  depth: number
}

const MAX_DEPTH = 64

const CHAR_SWITCHES = new Set(['lc', 'rc', 'bc', 'fc', 'vc'])

function run(text: string): Piece {
  if (text === '') return { omml: '', text: '' }
  return {
    omml:
      '<m:r><m:rPr><m:sty m:val="p"/></m:rPr>' +
      `<m:t xml:space="preserve">${escapeXmlText(text)}</m:t></m:r>`,
    text,
  }
}

function join(pieces: Piece[], sep = ''): Piece {
  return {
    omml: pieces.map((p) => p.omml).join(sep === '' ? '' : run(sep).omml),
    text: pieces.map((p) => p.text).join(sep),
  }
}

function slot(name: string, piece: Piece): string {
  return `<m:${name}>${piece.omml}</m:${name}>`
}

function val(name: string, v: string): string {
  return `<m:${name} m:val="${escapeXmlAttr(v)}"/>`
}

/** element sequence up to an unbalanced `)` or a top-level `,` (only inside args) */
function parseSeq(p: Parser, inArgs: boolean): Piece[] | null {
  const pieces: Piece[] = []
  let text = ''
  let depth = 0
  const flush = () => {
    if (text !== '') pieces.push(run(text))
    text = ''
  }
  while (p.i < p.src.length) {
    const ch = p.src[p.i]
    if (ch === '\\') {
      const next = p.src[p.i + 1] ?? ''
      if (/[A-Za-z]/.test(next)) {
        flush()
        const cmd = parseCommand(p)
        if (!cmd) return null
        pieces.push(cmd)
        continue
      }
      // escaped literal (`\,` `\(` `\)` `\\`)
      text += next
      p.i += 2
      continue
    }
    if (inArgs) {
      if (ch === '(') depth++
      else if (ch === ')') {
        if (depth === 0) break
        depth--
      } else if (ch === ',' && depth === 0) break
    }
    text += ch
    p.i++
  }
  flush()
  return pieces
}

function parseArgs(p: Parser): Piece[] | null {
  while (p.src[p.i] === ' ') p.i++
  if (p.src[p.i] !== '(') return null
  p.i++
  const args: Piece[] = []
  for (;;) {
    const seq = parseSeq(p, true)
    if (!seq) return null
    args.push(join(seq))
    const ch = p.src[p.i]
    if (ch === ',') {
      p.i++
      continue
    }
    if (ch === ')') {
      p.i++
      return args
    }
    return null
  }
}

function parseCommand(p: Parser): Piece | null {
  if (p.depth >= MAX_DEPTH) return null
  p.depth++
  try {
    return parseCommandInner(p)
  } finally {
    p.depth--
  }
}

function parseCommandInner(p: Parser): Piece | null {
  const cmd = p.src[p.i + 1].toLowerCase()
  p.i += 2
  const switches: Array<{ name: string; arg: string }> = []
  for (;;) {
    while (p.src[p.i] === ' ') p.i++
    const m = /^\\([A-Za-z]{2})(-?\d*)/.exec(p.src.slice(p.i))
    if (!m) break
    const name = m[1].toLowerCase()
    p.i += m[0].length
    let arg = m[2]
    if (CHAR_SWITCHES.has(name)) {
      if (p.src[p.i] !== '\\') return null
      arg = p.src[p.i + 1] ?? ''
      p.i += 2
    }
    switches.push({ name, arg })
  }
  const args = parseArgs(p)
  if (!args) return null
  const has = (name: string) => switches.some((s) => s.name === name)
  const argOf = (name: string) => switches.find((s) => s.name === name)?.arg
  switch (cmd) {
    case 'a': {
      // \co arrives from the file: a huge count fills each row's padding
      // loop (while cells.length < cols) and writes an invalid m:count.
      const rawCols = parseInt(argOf('co') ?? '1', 10)
      const cols = Number.isFinite(rawCols) ? Math.min(Math.max(1, rawCols), 64) : 1
      const jc = has('al') ? 'left' : has('ar') ? 'right' : 'center'
      const rows: string[] = []
      const lines: string[] = []
      for (let r = 0; r < args.length; r += cols) {
        const cells = args.slice(r, r + cols)
        while (cells.length < cols) cells.push(run(''))
        rows.push(`<m:mr>${cells.map((c) => slot('e', c)).join('')}</m:mr>`)
        lines.push(cells.map((c) => c.text).join(' '))
      }
      return {
        omml:
          '<m:m><m:mPr><m:mcs><m:mc><m:mcPr>' +
          val('count', String(cols)) +
          val('mcJc', jc) +
          `</m:mcPr></m:mc></m:mcs></m:mPr>${rows.join('')}</m:m>`,
        text: lines.join('\n'),
      }
    }
    case 'b': {
      const both = argOf('bc')
      const beg = both ?? argOf('lc') ?? '('
      const end = both ? closingOf(both) : (argOf('rc') ?? ')')
      const body = join(args, ',')
      return {
        omml:
          `<m:d><m:dPr>${val('begChr', beg)}${val('endChr', end)}</m:dPr>` +
          `${slot('e', body)}</m:d>`,
        text: `${beg}${body.text}${end}`,
      }
    }
    case 'f': {
      const [num = run(''), den = run('')] = args
      return {
        omml: `<m:f>${slot('num', num)}${slot('den', den)}</m:f>`,
        text: `${num.text}/${den.text}`,
      }
    }
    case 'i': {
      const chr = argOf('fc') ?? argOf('vc') ?? (has('su') ? '∑' : has('pr') ? '∏' : '∫')
      const [lo = run(''), hi = run(''), body = run('')] = args
      const limLoc = has('in') || chr === '∫' ? 'subSup' : 'undOvr'
      return {
        omml:
          `<m:nary><m:naryPr>${val('chr', chr)}${val('limLoc', limLoc)}` +
          (lo.text === '' ? val('subHide', '1') : '') +
          (hi.text === '' ? val('supHide', '1') : '') +
          `</m:naryPr>${slot('sub', lo)}${slot('sup', hi)}${slot('e', body)}</m:nary>`,
        text: `${chr}${lo.text}${hi.text} ${body.text}`,
      }
    }
    case 'l':
      return join(args, ',')
    case 'o':
      // overlay has no OMML counterpart; the pieces render side by side
      return join(args)
    case 'r': {
      if (args.length >= 2) {
        const [deg, body] = args
        return {
          omml: `<m:rad>${slot('deg', deg)}${slot('e', body)}</m:rad>`,
          text: `${deg.text}√(${body.text})`,
        }
      }
      const body = args[0] ?? run('')
      return {
        omml: `<m:rad><m:radPr>${val('degHide', '1')}</m:radPr><m:deg/>${slot('e', body)}</m:rad>`,
        text: `√(${body.text})`,
      }
    }
    case 's': {
      if (args.length === 1 && (has('up') || has('do'))) {
        const tag = has('up') ? 'sSup' : 'sSub'
        const script = has('up') ? 'sup' : 'sub'
        return {
          omml: `<m:${tag}><m:e/>${slot(script, args[0])}</m:${tag}>`,
          text: args[0].text,
        }
      }
      return {
        omml: `<m:eqArr>${args.map((a) => slot('e', a)).join('')}</m:eqArr>`,
        text: args.map((a) => a.text).join('\n'),
      }
    }
    case 'x': {
      const sides = { to: 'hideTop', bo: 'hideBot', le: 'hideLeft', ri: 'hideRight' } as const
      const named = Object.keys(sides).filter((k) => has(k))
      const hidden = named.length
        ? Object.entries(sides)
            .filter(([k]) => !named.includes(k))
            .map(([, tag]) => val(tag, '1'))
            .join('')
        : ''
      const body = join(args, ',')
      return {
        omml: `<m:borderBox><m:borderBoxPr>${hidden}</m:borderBoxPr>${slot('e', body)}</m:borderBox>`,
        text: body.text,
      }
    }
    default:
      return null
  }
}

function closingOf(open: string): string {
  return { '(': ')', '[': ']', '{': '}', '<': '>', '⟨': '⟩' }[open] ?? open
}

/** null when the instruction is not an EQ field or uses an unsupported switch */
export function eqFieldToOmml(instr: string): EqField | null {
  const m = /^\s*EQ\b\s*([\s\S]*?)\s*$/i.exec(instr)
  if (!m) return null
  // general switches (\* MERGEFORMAT ...) are not part of the layout
  const p: Parser = { src: m[1].replace(/\\\*\s*[A-Za-z]+/g, '').trim(), i: 0, depth: 0 }
  const seq = parseSeq(p, false)
  if (!seq || p.i < p.src.length) return null
  const body = join(seq)
  if (body.omml === '') return null
  return { omml: `<m:oMath>${body.omml}</m:oMath>`, text: body.text }
}

const FLD_CHAR_RE =
  /<w:fldChar[^>]*w:fldCharType=(?:"(begin|separate|end)"|'(begin|separate|end)')/g

/**
 * Paragraph XML with every renderable EQ field replaced by a plain run of its
 * linear text (in the field code's rPr), so the field-result text/size scans
 * see the layout Word would compute.
 */
export function inlineEqFieldResults(xml: string): string {
  if (!/<w:instrText[^>]*>\s*EQ\b/i.test(xml)) return xml
  let out = ''
  let cursor = 0
  let depth = 0
  let spanStart = -1
  let m: RegExpExecArray | null
  FLD_CHAR_RE.lastIndex = 0
  while ((m = FLD_CHAR_RE.exec(xml)) !== null) {
    const kind = m[1] ?? m[2]
    if (kind === 'begin') {
      if (depth === 0) spanStart = runStartBefore(xml, m.index)
      depth++
    } else if (kind === 'end') {
      depth = Math.max(0, depth - 1)
      if (depth !== 0 || spanStart < 0) continue
      const spanEnd = xml.indexOf('</w:r>', m.index) + '</w:r>'.length
      const span = xml.slice(spanStart, spanEnd)
      const instr = decodeEntities(
        Array.from(span.matchAll(/<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/g), (x) => x[1]).join(
          '',
        ),
      )
      const eq = eqFieldToOmml(instr)
      if (eq) {
        const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(span)?.[0] ?? ''
        out +=
          xml.slice(cursor, spanStart) +
          `<w:r>${rPr}<w:t xml:space="preserve">${escapeXmlText(eq.text)}</w:t></w:r>`
        cursor = spanEnd
      }
      spanStart = -1
    }
  }
  return out + xml.slice(cursor)
}

function runStartBefore(xml: string, index: number): number {
  const a = xml.lastIndexOf('<w:r>', index)
  const b = xml.lastIndexOf('<w:r ', index)
  return Math.max(a, b)
}

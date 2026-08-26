/**
 * Budgeted regular-expression engine for layout scripts. Native `RegExp`
 * matching runs outside the interpreter's step accounting, so a
 * catastrophic-backtracking pattern could freeze the renderer past the step
 * limit. This matcher supports the subset layout scripts need (literals,
 * classes, groups, alternation, quantifiers, anchors, i/m/s flags) and charges
 * every match step against a hard budget instead.
 */

type PredefName = 'd' | 'D' | 'w' | 'W' | 's' | 'S'

type RxNode =
  | { kind: 'seq'; parts: RxNode[] }
  | { kind: 'alt'; options: RxNode[] }
  | { kind: 'char'; ch: string }
  | { kind: 'dot' }
  | {
      kind: 'class'
      negated: boolean
      singles: number[]
      ranges: [number, number][]
      predefs: PredefName[]
    }
  | { kind: 'predef'; name: PredefName }
  | { kind: 'anchor'; at: 'start' | 'end' | 'word' | 'nonWord' }
  | { kind: 'repeat'; min: number; max: number; inner: RxNode }

const MAX_PATTERN_LENGTH = 500
const MAX_REPEAT_COUNT = 1000
const MAX_MATCH_STEPS = 1_000_000
const MAX_MATCH_DEPTH = 2000

const BUDGET_MESSAGE = 'Layout script regular expression exceeded its execution budget'

export interface BoundedRegex {
  test(input: string): boolean
}

const unavailable = (feature: string): never => {
  throw new Error(`Regular-expression ${feature} is not available in layout scripts`)
}

const isLineTerminator = (c: string): boolean =>
  c === '\n' || c === '\r' || c === '\u2028' || c === '\u2029'

const predefMatch = (name: PredefName, c: string): boolean => {
  switch (name) {
    case 'd':
      return c >= '0' && c <= '9'
    case 'D':
      return !(c >= '0' && c <= '9')
    case 'w':
      return /[0-9A-Za-z_]/.test(c)
    case 'W':
      return !/[0-9A-Za-z_]/.test(c)
    case 's':
      return /\s/.test(c)
    case 'S':
      return !/\s/.test(c)
  }
}

export function compileBoundedRegex(source: string, flags: string): BoundedRegex {
  for (const flag of flags) if (!'gimsd'.includes(flag)) unavailable(`flag "${flag}"`)
  if (source.length > MAX_PATTERN_LENGTH)
    throw new Error(
      `Regular-expression pattern is longer than the ${MAX_PATTERN_LENGTH}-character layout-script limit`,
    )
  const ignoreCase = flags.includes('i')
  const multiline = flags.includes('m')
  const dotAll = flags.includes('s')

  let pos = 0
  const syntaxError = (): never => {
    throw new Error(`Invalid regular expression: /${source}/${flags}`)
  }
  const charNode = (ch: string): RxNode => ({ kind: 'char', ch })

  const parseEscape = (inClass: boolean): RxNode => {
    pos += 1
    const ch = source[pos]
    if (ch === undefined) return syntaxError()
    pos += 1
    switch (ch) {
      case 'd':
      case 'D':
      case 'w':
      case 'W':
      case 's':
      case 'S':
        return { kind: 'predef', name: ch }
      case 'b':
        return inClass ? charNode('\b') : { kind: 'anchor', at: 'word' }
      case 'B':
        if (inClass) syntaxError()
        return { kind: 'anchor', at: 'nonWord' }
      case 'n':
        return charNode('\n')
      case 'r':
        return charNode('\r')
      case 't':
        return charNode('\t')
      case 'f':
        return charNode('\f')
      case 'v':
        return charNode('\v')
      case '0':
        return charNode('\0')
      case 'x': {
        const hex = source.slice(pos, pos + 2)
        if (!/^[0-9a-fA-F]{2}$/.test(hex)) syntaxError()
        pos += 2
        return charNode(String.fromCharCode(parseInt(hex, 16)))
      }
      case 'u': {
        const hex = source.slice(pos, pos + 4)
        if (!/^[0-9a-fA-F]{4}$/.test(hex)) syntaxError()
        pos += 4
        return charNode(String.fromCharCode(parseInt(hex, 16)))
      }
      case 'c': {
        const letter = source[pos]
        if (!letter || !/[a-zA-Z]/.test(letter)) return syntaxError()
        pos += 1
        return charNode(String.fromCharCode(letter.toUpperCase().charCodeAt(0) - 64))
      }
      case 'k':
        return unavailable('named backreference')
      case 'p':
      case 'P':
        return unavailable('unicode property escape')
      default:
        if (ch >= '1' && ch <= '9') return unavailable('backreference')
        if (/[a-zA-Z]/.test(ch)) return syntaxError()
        return charNode(ch)
    }
  }

  const parseClass = (): RxNode => {
    pos += 1
    let negated = false
    if (source[pos] === '^') {
      negated = true
      pos += 1
    }
    const singles: number[] = []
    const ranges: [number, number][] = []
    const predefs: PredefName[] = []
    while (source[pos] !== ']') {
      if (pos >= source.length) syntaxError()
      let lo: number
      if (source[pos] === '\\') {
        const node = parseEscape(true)
        if (node.kind === 'predef') {
          predefs.push(node.name)
          continue
        }
        lo = (node as { ch: string }).ch.charCodeAt(0)
      } else {
        lo = source.charCodeAt(pos)
        pos += 1
      }
      if (source[pos] === '-' && pos + 1 < source.length && source[pos + 1] !== ']') {
        pos += 1
        let hi: number
        if (source[pos] === '\\') {
          const node = parseEscape(true)
          if (node.kind === 'predef') return syntaxError()
          hi = (node as { ch: string }).ch.charCodeAt(0)
        } else {
          hi = source.charCodeAt(pos)
          pos += 1
        }
        if (hi < lo) syntaxError()
        ranges.push([lo, hi])
      } else {
        singles.push(lo)
      }
    }
    pos += 1
    return { kind: 'class', negated, singles, ranges, predefs }
  }

  const parseGroup = (): RxNode => {
    pos += 1
    if (source[pos] === '?') {
      const next = source[pos + 1]
      if (next === ':') pos += 2
      else if (next === '<' && source[pos + 2] !== '=' && source[pos + 2] !== '!') {
        const end = source.indexOf('>', pos + 2)
        if (end === -1) return syntaxError()
        pos = end + 1
      } else return unavailable(next === '=' || next === '!' ? 'lookahead' : 'lookbehind')
    }
    const inner = parseAlternation()
    if (source[pos] !== ')') syntaxError()
    pos += 1
    return inner
  }

  const parseAtom = (): RxNode => {
    const ch = source[pos]
    if (ch === undefined) return syntaxError()
    switch (ch) {
      case '^':
        pos += 1
        return { kind: 'anchor', at: 'start' }
      case '$':
        pos += 1
        return { kind: 'anchor', at: 'end' }
      case '.':
        pos += 1
        return { kind: 'dot' }
      case '(':
        return parseGroup()
      case '[':
        return parseClass()
      case '\\':
        return parseEscape(false)
      case ')':
      case '*':
      case '+':
      case '?':
        return syntaxError()
      default:
        pos += 1
        return charNode(ch)
    }
  }

  const parseQuantifier = (): { min: number; max: number } | undefined => {
    const ch = source[pos]
    let quantifier: { min: number; max: number } | undefined
    if (ch === '*') {
      quantifier = { min: 0, max: Infinity }
      pos += 1
    } else if (ch === '+') {
      quantifier = { min: 1, max: Infinity }
      pos += 1
    } else if (ch === '?') {
      quantifier = { min: 0, max: 1 }
      pos += 1
    } else if (ch === '{') {
      const match = /^\{(\d+)(?:,(\d*))?\}/.exec(source.slice(pos))
      if (match) {
        const min = Number(match[1])
        const max = match[2] === undefined ? min : match[2] === '' ? Infinity : Number(match[2])
        if (max < min) syntaxError()
        if (min > MAX_REPEAT_COUNT || (max !== Infinity && max > MAX_REPEAT_COUNT))
          unavailable(`repetition beyond {${MAX_REPEAT_COUNT}}`)
        quantifier = { min, max }
        pos += match[0].length
      }
    }
    if (!quantifier) return undefined
    if (source[pos] === '?') pos += 1
    if (source[pos] === '*' || source[pos] === '+' || source[pos] === '?') syntaxError()
    return quantifier
  }

  const parseQuantified = (): RxNode => {
    const atom = parseAtom()
    const quantifier = parseQuantifier()
    if (!quantifier) return atom
    if (atom.kind === 'anchor') syntaxError()
    return { kind: 'repeat', min: quantifier.min, max: quantifier.max, inner: atom }
  }

  const parseSequence = (): RxNode => {
    const parts: RxNode[] = []
    while (pos < source.length && source[pos] !== '|' && source[pos] !== ')')
      parts.push(parseQuantified())
    return parts.length === 1 ? parts[0] : { kind: 'seq', parts }
  }

  const parseAlternation = (): RxNode => {
    const options = [parseSequence()]
    while (source[pos] === '|') {
      pos += 1
      options.push(parseSequence())
    }
    return options.length === 1 ? options[0] : { kind: 'alt', options }
  }

  const root = parseAlternation()
  if (pos < source.length) syntaxError()

  const charEq = (a: string, b: string): boolean =>
    a === b || (ignoreCase && a.toLowerCase() === b.toLowerCase())

  const classMatch = (node: Extract<RxNode, { kind: 'class' }>, c: string): boolean => {
    const codeMatches = (code: number): boolean =>
      node.singles.includes(code) || node.ranges.some(([lo, hi]) => code >= lo && code <= hi)
    let hit = codeMatches(c.charCodeAt(0)) || node.predefs.some((name) => predefMatch(name, c))
    if (!hit && ignoreCase) {
      const lower = c.toLowerCase()
      const upper = c.toUpperCase()
      hit =
        (lower !== c && lower.length === 1 && codeMatches(lower.charCodeAt(0))) ||
        (upper !== c && upper.length === 1 && codeMatches(upper.charCodeAt(0)))
    }
    return node.negated ? !hit : hit
  }

  const singleMatch = (node: RxNode, c: string): boolean => {
    switch (node.kind) {
      case 'char':
        return charEq(c, node.ch)
      case 'dot':
        return dotAll || !isLineTerminator(c)
      case 'predef':
        return predefMatch(node.name, c)
      case 'class':
        return classMatch(node, c)
      default:
        return false
    }
  }

  const isSingle = (node: RxNode): boolean =>
    node.kind === 'char' || node.kind === 'dot' || node.kind === 'predef' || node.kind === 'class'

  const isWordChar = (c: string | undefined): boolean => c !== undefined && /[0-9A-Za-z_]/.test(c)

  const startsAnchored = (node: RxNode): boolean => {
    switch (node.kind) {
      case 'anchor':
        return node.at === 'start'
      case 'seq':
        return node.parts.length > 0 && startsAnchored(node.parts[0])
      case 'alt':
        return node.options.every(startsAnchored)
      case 'repeat':
        return node.min >= 1 && startsAnchored(node.inner)
      default:
        return false
    }
  }
  const anchoredOnly = !multiline && startsAnchored(root)

  const test = (input: string): boolean => {
    let budget = MAX_MATCH_STEPS
    let depth = 0
    const step = (): void => {
      budget -= 1
      if (budget < 0) throw new Error(BUDGET_MESSAGE)
    }

    const match = (node: RxNode, at: number, cont: (p: number) => boolean): boolean => {
      step()
      depth += 1
      if (depth > MAX_MATCH_DEPTH) {
        depth -= 1
        throw new Error(BUDGET_MESSAGE)
      }
      try {
        switch (node.kind) {
          case 'char':
          case 'dot':
          case 'predef':
          case 'class':
            return at < input.length && singleMatch(node, input[at]) && cont(at + 1)
          case 'anchor': {
            let ok: boolean
            switch (node.at) {
              case 'start':
                ok = at === 0 || (multiline && isLineTerminator(input[at - 1]))
                break
              case 'end':
                ok = at === input.length || (multiline && isLineTerminator(input[at]))
                break
              case 'word':
                ok = isWordChar(input[at - 1]) !== isWordChar(input[at])
                break
              case 'nonWord':
                ok = isWordChar(input[at - 1]) === isWordChar(input[at])
                break
            }
            return ok && cont(at)
          }
          case 'seq': {
            const run = (index: number, p: number): boolean =>
              index >= node.parts.length
                ? cont(p)
                : match(node.parts[index], p, (np) => run(index + 1, np))
            return run(0, at)
          }
          case 'alt':
            return node.options.some((option) => match(option, at, cont))
          case 'repeat': {
            if (isSingle(node.inner)) {
              let count = 0
              let p = at
              while (count < node.max && p < input.length && singleMatch(node.inner, input[p])) {
                step()
                count += 1
                p += 1
              }
              while (count >= node.min) {
                step()
                if (cont(p)) return true
                count -= 1
                p -= 1
              }
              return false
            }
            const rep = (count: number, p: number): boolean => {
              step()
              if (
                count < node.max &&
                match(node.inner, p, (np) => (np === p ? cont(p) : rep(count + 1, np)))
              )
                return true
              return count >= node.min && cont(p)
            }
            return rep(0, at)
          }
        }
      } finally {
        depth -= 1
      }
    }

    const lastStart = anchoredOnly ? 0 : input.length
    for (let start = 0; start <= lastStart; start += 1) {
      step()
      if (match(root, start, () => true)) return true
    }
    return false
  }

  return { test }
}

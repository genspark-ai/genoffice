/// Formula model for pivot calculated fields: a small expression language of
/// basic arithmetic + parentheses + field references, parsed by recursive descent
/// into an AST and evaluated safely (no eval). Field references may be wrapped in
/// single quotes or written as bare names; aligned with the syntax subset of
/// Excel calculated fields' cacheField@formula for easy OOXML round-trips.

export class PivotFormulaError extends Error {}

export type PivotFormulaAst =
  | { readonly t: 'num'; readonly value: number }
  | { readonly t: 'field'; readonly name: string }
  | { readonly t: 'neg'; readonly operand: PivotFormulaAst }
  | {
      readonly t: 'bin'
      readonly op: '+' | '-' | '*' | '/'
      readonly left: PivotFormulaAst
      readonly right: PivotFormulaAst
    }

interface Token {
  readonly kind: 'num' | 'field' | 'op' | 'lparen' | 'rparen'
  readonly text: string
}

/// Characters allowed in bare field names: anything other than letters (incl.
/// CJK)/digits/underscore/dot/space is treated as an operator boundary; spaces
/// only separate and never enter a bare name.
const BARE_NAME_CHAR = /[^\s+\-*/()']/

function tokenize(formula: string): Token[] {
  const tokens: Token[] = []
  let at = 0
  while (at < formula.length) {
    const ch = formula[at]!
    if (/\s/.test(ch)) {
      at += 1
      continue
    }
    if ('+-*/'.includes(ch)) {
      tokens.push({ kind: 'op', text: ch })
      at += 1
      continue
    }
    if (ch === '(') {
      tokens.push({ kind: 'lparen', text: ch })
      at += 1
      continue
    }
    if (ch === ')') {
      tokens.push({ kind: 'rparen', text: ch })
      at += 1
      continue
    }
    if (ch === "'") {
      // Quoted field name; '' escapes a single quote.
      let name = ''
      at += 1
      for (;;) {
        if (at >= formula.length) throw new PivotFormulaError('Unclosed quote in the formula.')
        if (formula[at] === "'") {
          if (formula[at + 1] === "'") {
            name += "'"
            at += 2
            continue
          }
          at += 1
          break
        }
        name += formula[at]
        at += 1
      }
      if (name.trim() === '')
        throw new PivotFormulaError('The formula references an empty field name.')
      tokens.push({ kind: 'field', text: name.trim() })
      continue
    }
    if (/[0-9.]/.test(ch)) {
      const match = /^\d*\.?\d+(?:[eE][+-]?\d+)?/.exec(formula.slice(at))
      if (match && Number.isFinite(Number(match[0]))) {
        tokens.push({ kind: 'num', text: match[0] })
        at += match[0].length
        continue
      }
      throw new PivotFormulaError(
        `Cannot parse a number in the formula: "${formula.slice(at, at + 8)}"`,
      )
    }
    if (BARE_NAME_CHAR.test(ch)) {
      let name = ''
      while (at < formula.length && BARE_NAME_CHAR.test(formula[at]!)) {
        name += formula[at]
        at += 1
      }
      tokens.push({ kind: 'field', text: name })
      continue
    }
    throw new PivotFormulaError(`Unrecognized character in the formula: "${ch}"`)
  }
  return tokens
}

/// Parses a formula (optional leading =). fieldNames are the referenceable field
/// names (case-insensitive; references are normalized to the original names
/// passed in); referencing an unknown field throws, failing closed.
export function parsePivotFormula(formula: string, fieldNames: readonly string[]): PivotFormulaAst {
  const canonicalByKey = new Map<string, string>()
  for (const name of fieldNames) canonicalByKey.set(name.trim().toLowerCase(), name)
  const tokens = tokenize(formula.replace(/^\s*=/, ''))
  let at = 0

  const peek = (): Token | undefined => tokens[at]
  const next = (): Token => {
    const token = tokens[at]
    if (!token) throw new PivotFormulaError('The formula ended unexpectedly.')
    at += 1
    return token
  }

  // expr := term (('+'|'-') term)* ; term := unary (('*'|'/') unary)* ;
  // unary := '-' unary | primary ; primary := num | field | '(' expr ')'
  const parseExpr = (): PivotFormulaAst => {
    let left = parseTerm()
    while (peek()?.kind === 'op' && (peek()!.text === '+' || peek()!.text === '-')) {
      const op = next().text as '+' | '-'
      left = { t: 'bin', op, left, right: parseTerm() }
    }
    return left
  }
  const parseTerm = (): PivotFormulaAst => {
    let left = parseUnary()
    while (peek()?.kind === 'op' && (peek()!.text === '*' || peek()!.text === '/')) {
      const op = next().text as '*' | '/'
      left = { t: 'bin', op, left, right: parseUnary() }
    }
    return left
  }
  const parseUnary = (): PivotFormulaAst => {
    if (peek()?.kind === 'op' && peek()!.text === '-') {
      next()
      return { t: 'neg', operand: parseUnary() }
    }
    if (peek()?.kind === 'op' && peek()!.text === '+') {
      next()
      return parseUnary()
    }
    return parsePrimary()
  }
  const parsePrimary = (): PivotFormulaAst => {
    const token = next()
    if (token.kind === 'num') return { t: 'num', value: Number(token.text) }
    if (token.kind === 'field') {
      const canonical = canonicalByKey.get(token.text.trim().toLowerCase())
      if (canonical === undefined) {
        throw new PivotFormulaError(`The formula references a nonexistent field "${token.text}".`)
      }
      return { t: 'field', name: canonical }
    }
    if (token.kind === 'lparen') {
      const inner = parseExpr()
      if (next().kind !== 'rparen')
        throw new PivotFormulaError('Unclosed parenthesis in the formula.')
      return inner
    }
    throw new PivotFormulaError(`Unexpected "${token.text}" in the formula.`)
  }

  const ast = parseExpr()
  if (at !== tokens.length) {
    throw new PivotFormulaError(`The formula cannot be parsed further at "${tokens[at]!.text}".`)
  }
  return ast
}

/// Evaluation: field references resolve through valueOf (under calculated-field
/// semantics, the group's aggregated field value; empty groups count as 0);
/// division by 0 returns null (the caller blanks the cell).
export function evaluatePivotFormula(
  ast: PivotFormulaAst,
  valueOf: (fieldName: string) => number | null,
): number | null {
  switch (ast.t) {
    case 'num':
      return ast.value
    case 'field':
      return valueOf(ast.name) ?? 0
    case 'neg': {
      const operand = evaluatePivotFormula(ast.operand, valueOf)
      return operand === null ? null : -operand
    }
    case 'bin': {
      const left = evaluatePivotFormula(ast.left, valueOf)
      const right = evaluatePivotFormula(ast.right, valueOf)
      if (left === null || right === null) return null
      switch (ast.op) {
        case '+':
          return left + right
        case '-':
          return left - right
        case '*':
          return left * right
        case '/':
          return right === 0 ? null : left / right
      }
    }
  }
}

/// Normalized output: field references are always single-quoted (inner quotes
/// escaped as ''), compatible with Excel's cacheField@formula notation; all
/// operations are explicitly parenthesized to avoid precedence ambiguity.
export function formatPivotFormula(ast: PivotFormulaAst): string {
  switch (ast.t) {
    case 'num':
      return String(ast.value)
    case 'field':
      return `'${ast.name.replace(/'/g, "''")}'`
    case 'neg':
      return `(-${formatPivotFormula(ast.operand)})`
    case 'bin':
      return `(${formatPivotFormula(ast.left)}${ast.op}${formatPivotFormula(ast.right)})`
  }
}

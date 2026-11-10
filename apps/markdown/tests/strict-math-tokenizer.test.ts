import { describe, expect, it } from 'vitest'
import type { MarkdownLexerConfiguration, MarkdownTokenizer } from '@tiptap/core'
import { buildMathExtensions } from '../src/renderer/editor/math'

const tokenizer = (() => {
  const extension = buildMathExtensions().find((e) => e.name === 'inlineMath')!
  return extension.config.markdownTokenizer as Required<MarkdownTokenizer>
})()

function tokenize(src: string) {
  const start = typeof tokenizer.start === 'function' ? tokenizer.start(src) : src.indexOf('$')
  if (start < 0) return null
  return tokenizer.tokenize(src.slice(start), [], {} as MarkdownLexerConfiguration)
}

function latexIn(src: string): string | null {
  const token = tokenize(src) as { latex?: string } | null | undefined
  return token?.latex ?? null
}

function startIn(src: string): number {
  return typeof tokenizer.start === 'function' ? tokenizer.start(src) : src.indexOf('$')
}

describe('strict inline math tokenizer', () => {
  it('never opens math on an escaped dollar', () => {
    expect(latexIn(String.raw`\$5\$`)).toBeNull()
    expect(latexIn(String.raw`\$\alpha$`)).toBeNull()
    expect(latexIn(String.raw`\$x^2$`)).toBeNull()
    expect(latexIn(String.raw`costs \$5 and \$10.`)).toBeNull()
  })

  it('does not let an escaped dollar close math either', () => {
    expect(latexIn(String.raw`$5\$`)).toBeNull()
    expect(latexIn(String.raw`$x^2\$`)).toBeNull()
    expect(latexIn(String.raw`$a\$b`)).toBeNull()
  })

  it('leaves a formula with an escaped dollar literal rather than half-matching it', () => {
    expect(latexIn(String.raw`$a\$b$`)).toBeNull()
  })

  it('still opens math after an escaped backslash', () => {
    expect(latexIn(String.raw`\\$x^2$`)).toBe('x^2')
    expect(latexIn(String.raw`a \\$b$ c`)).toBe('b')
    expect(latexIn(String.raw`\\\\$c$`)).toBe('c')
  })

  it('keeps plain currency literal', () => {
    expect(latexIn('Price $5. Total $10.')).toBeNull()
    expect(latexIn('I paid $5 and $10 in total')).toBeNull()
    expect(latexIn('costs $5 today')).toBeNull()
  })

  it('rejects a closing dollar followed by a digit or another dollar', () => {
    expect(latexIn('$x$5')).toBeNull()
    expect(latexIn('$x$$')).toBeNull()
  })

  it('does not treat a display delimiter as inline math', () => {
    expect(latexIn('$$x$$')).toBeNull()
  })

  it('parses ordinary inline math', () => {
    expect(latexIn('a $x^2$ b')).toBe('x^2')
    expect(latexIn('the variable $x_{1}$ is free')).toBe('x_{1}')
    expect(latexIn('a $b$ c $d$ e')).toBe('b')
  })

  it('reports the opening dollar position for an unescaped delimiter', () => {
    expect(startIn('$x$')).toBe(0)
    expect(startIn('a $x$')).toBe(2)
    expect(startIn(String.raw`\$x$`)).toBe(3)
    expect(startIn(String.raw`\$x\$`)).toBe(-1)
    expect(startIn('no dollars here')).toBe(-1)
  })
})

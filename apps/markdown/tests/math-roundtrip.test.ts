import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { getDefaults, marked } from 'marked'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { containsMathSyntax, matchInlineMath } from '../src/renderer/editor/mathSyntax'

const editors: Editor[] = []
afterEach(() => {
  for (const e of editors.splice(0)) e.destroy()
})

function open(md: string): Editor {
  marked.setOptions(getDefaults())
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
    }),
    content: '',
  })
  editors.push(editor)
  editor.commands.setContent(md, { contentType: 'markdown' })
  return editor
}

const mathCount = (e: Editor) =>
  (JSON.stringify(e.getJSON()).match(/"type":"inlineMath"/g) ?? []).length

describe('inline math vs currency', () => {
  it('keeps escaped dollars escaped on save so they stay text on reopen', () => {
    const md = 'a \\$b\\$ c'
    const first = open(md)
    expect(mathCount(first)).toBe(0)
    const saved = first.getMarkdown()
    expect(saved).toBe(md)
    expect(mathCount(open(saved))).toBe(0)
    // no longer math-shaped, so the escape may drop, but the text must survive a reopen
    const currency = open('paid \\$5 and 10\\$ today').getMarkdown()
    expect(mathCount(open(currency))).toBe(0)
  })

  it('does not read a suffix amount as the closing delimiter', () => {
    const md = 'costs $5 and 10$ more'
    const editor = open(md)
    expect(mathCount(editor)).toBe(0)
    expect(editor.getMarkdown()).toBe(md)
  })

  it('leaves plain currency unescaped and formulas intact', () => {
    for (const md of ['paid $5 and $10 today', 'from $5 to $.50 cents', 'US$5 and US$10']) {
      const editor = open(md)
      expect(mathCount(editor)).toBe(0)
      expect(editor.getMarkdown()).toBe(md)
    }
    const math = open('formula $E=mc^2$ ok')
    expect(mathCount(math)).toBe(1)
    expect(math.getMarkdown()).toBe('formula $E=mc^2$ ok')
  })

  it('matchInlineMath / containsMathSyntax', () => {
    expect(matchInlineMath('$x^2$ y')?.latex).toBe('x^2')
    expect(matchInlineMath('$5 and 10$ more')).toBeUndefined()
    expect(matchInlineMath('$5 + x$')?.latex).toBe('5 + x')
    expect(matchInlineMath('$2 + 3$')?.latex).toBe('2 + 3')
    expect(matchInlineMath('$1 + 1 = 2$')?.latex).toBe('1 + 1 = 2')
    expect(matchInlineMath('$3 \\times 4$')?.latex).toBe('3 \\times 4')
    expect(matchInlineMath('$5 to 10$')).toBeUndefined()
    expect(containsMathSyntax('sum $2 + 3$')).toBe(true)
    expect(containsMathSyntax('paid $5 and $10')).toBe(false)
    expect(containsMathSyntax('a $b$ c')).toBe(true)
    expect(containsMathSyntax('a $$b$$ c')).toBe(true)
  })
})

/**
 * `==highlight==` round trip: parse to the highlight mark, save back as the
 * same syntax, and survive the edge cases a delimiter scheme invites.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { getDefaults, marked } from 'marked'
import { buildExtensions } from '../src/renderer/editor/extensions'

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

const hlCount = (e: Editor) =>
  (JSON.stringify(e.getJSON()).match(/"type":"highlight"/g) ?? []).length

describe('==highlight== mark', () => {
  it('parses and saves back the same syntax', () => {
    const editor = open('before ==marked text== after')
    expect(hlCount(editor)).toBe(1)
    expect(editor.getMarkdown()).toBe('before ==marked text== after')
  })

  it('keeps two highlights on one line distinct', () => {
    const editor = open('==a== and ==b==')
    expect(hlCount(editor)).toBe(2)
    expect(editor.getMarkdown()).toBe('==a== and ==b==')
  })

  it('treats an unclosed or spaced-out run as plain text', () => {
    for (const md of ['a == b', 'a ==unclosed', '== starts with space']) {
      const editor = open(md)
      expect(hlCount(editor)).toBe(0)
      expect(editor.getMarkdown()).toBe(md)
    }
  })

  it('keeps highlight and bold through nested-delimiter reordering', () => {
    // delimiter marks nest by tree order, so the serializer may swap the
    // wrapping (`**==x==**`) — the model is order-free and one save settles it
    const editor = open('==**bold inside**==')
    expect(hlCount(editor)).toBe(1)
    const saved = editor.getMarkdown()
    expect(hlCount(open(saved))).toBe(1)
    expect(open(saved).getMarkdown()).toBe(saved)
  })

  it('keeps text beside a nested mark inside the run, and text around it outside', () => {
    for (const md of ['==hello **bold**==', 'see ==**x**== here', 'see ==a **b** c== here']) {
      // the mark sits on each text node it spans, so count the text instead
      const editor = open(md)
      const saved = editor.getMarkdown()
      expect(saved.replace(/\*\*/g, '')).toBe(md.replace(/\*\*/g, ''))
      expect(hlCount(open(saved))).toBe(hlCount(editor))
      expect(hlCount(editor)).toBeGreaterThan(0)
    }
  })

  it('keeps plain equations like a == b unhighlighted across saves', () => {
    for (const md of ['if a == b then c', 'if a == **b** == c', 'x == *y* and z == w']) {
      const editor = open(md)
      expect(hlCount(editor)).toBe(0)
      expect(editor.getMarkdown()).toBe(md)
    }
  })
})

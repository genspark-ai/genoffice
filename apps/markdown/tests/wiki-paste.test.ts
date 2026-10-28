/**
 * Pasted wiki syntax. The tokenizers only run on load, so a `[[link]]` or
 * `![[embed]]` that arrived through the clipboard used to sit as plain text
 * until the file was reopened; after a paste it now takes the shape the
 * tokenizers would have built. Typing is not a paste and is left to the `[[`
 * suggester.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { collectWikiTargets } from '../src/renderer/editor/wikilink'
import { collectWikiEmbedTargets } from '../src/renderer/editor/wikiEmbed'

const editors: Editor[] = []
afterAll(() => {
  for (const e of editors) e.destroy()
})

function openDoc(source: string): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: {
        onOpen: () => {},
        onUpdate: () => {},
        onKeyDown: () => false,
        onClose: () => {},
      },
      slashItems: () => [],
    }),
    content: '',
  })
  editors.push(editor)
  editor.commands.setContent(source, { contentType: 'markdown' })
  return editor
}

/** the caret at the end of the first textblock containing `text` */
function caretAfter(editor: Editor, text: string): void {
  let at: number | null = null
  editor.state.doc.descendants((node, pos) => {
    if (at !== null) return false
    if (node.isTextblock && node.textContent.includes(text)) at = pos + node.nodeSize - 1
    return at === null
  })
  if (at === null) throw new Error(`no block containing ${text}`)
  editor.commands.setTextSelection(at)
}

/** jsdom has no ClipboardEvent; the view only needs an event to hand to the paste props */
function paste(editor: Editor, text: string): void {
  editor.view.pasteText(text, new Event('paste') as unknown as ClipboardEvent)
}

/** every text run with its marks, and every inline node, in document order */
function inlines(editor: Editor): string[] {
  const out: string[] = []
  editor.state.doc.descendants((node) => {
    if (node.isText) out.push(`${node.text}:${node.marks.map((m) => m.type.name).join(',')}`)
    else if (node.isInline) out.push(`<${node.type.name} ${String(node.attrs.raw ?? '')}>`)
  })
  return out
}

describe('pasted wiki syntax goes live', () => {
  it('marks a pasted [[link]] so the resolver sees it', () => {
    const editor = openDoc('Start.')
    caretAfter(editor, 'Start')
    paste(editor, ' see [[Other#^abc123]] now')
    expect(inlines(editor)).toEqual(['Start. see :', '[[Other#^abc123]]:wikiLink', ' now:'])
    expect(collectWikiTargets(editor.state.doc)).toEqual(['[[Other#^abc123]]'])
    // the file reads exactly as pasted
    expect(editor.getMarkdown()).toBe('Start. see [[Other#^abc123]] now')
  })

  it('turns a pasted ![[embed]] into the embed node', () => {
    const editor = openDoc('Start.')
    caretAfter(editor, 'Start')
    paste(editor, ' ![[Pic.png|200]] tail')
    expect(inlines(editor)).toEqual(['Start. :', '<wikiEmbed ![[Pic.png|200]]>', ' tail:'])
    expect(collectWikiEmbedTargets(editor.state.doc)).toEqual(['Pic.png|200'])
    expect(editor.getMarkdown()).toBe('Start. ![[Pic.png|200]] tail')
  })

  it('handles several literals and a link in a later pasted paragraph', () => {
    const editor = openDoc('Start.')
    caretAfter(editor, 'Start')
    paste(editor, ' [[A]] and [[B|bee]]\n\nSecond [[C]].')
    expect(collectWikiTargets(editor.state.doc)).toEqual(['[[A]]', '[[B|bee]]', '[[C]]'])
  })

  it('leaves an escaped \\[[literal]] and a bare [[ alone', () => {
    const editor = openDoc('Start.')
    caretAfter(editor, 'Start')
    paste(editor, ' \\[[not a link]] and [[ open')
    expect(collectWikiTargets(editor.state.doc)).toEqual([])
    expect(inlines(editor).every((run) => !run.includes('wikiLink'))).toBe(true)
  })

  it('does not touch a code block or a code span', () => {
    const block = openDoc('```\ncode\n```')
    caretAfter(block, 'code')
    paste(block, ' [[x]]')
    expect(collectWikiTargets(block.state.doc)).toEqual([])

    const span = openDoc('a `code` b')
    // inside the code span: the pasted text carries the code mark
    let inside: number | null = null
    span.state.doc.descendants((node, pos) => {
      if (inside === null && node.isText && node.marks.some((m) => m.type.name === 'code'))
        inside = pos + 2
    })
    span.commands.setTextSelection(inside!)
    paste(span, '[[y]]')
    expect(collectWikiTargets(span.state.doc)).toEqual([])
    expect(inlines(span).some((run) => run.startsWith('co[[y]]de:code'))).toBe(true)
  })

  it('is a paste rule, not a typing rule', () => {
    const editor = openDoc('Start.')
    caretAfter(editor, 'Start')
    editor.commands.insertContent(' [[Typed]]')
    expect(collectWikiTargets(editor.state.doc)).toEqual([])
  })

  it('leaves a link that is already live as it is', () => {
    const editor = openDoc('See [[Existing]] here.')
    caretAfter(editor, 'here')
    paste(editor, ' and [[New]]')
    expect(collectWikiTargets(editor.state.doc)).toEqual(['[[Existing]]', '[[New]]'])
    expect(editor.getMarkdown()).toBe('See [[Existing]] here. and [[New]]')
  })
})

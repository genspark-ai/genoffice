/**
 * Story view (editor/story-view.ts): the margin markers and the typing rule.
 *
 * The view itself is CSS (styles.css `.view-story`); what needs a test here
 * are the two things the CSS cannot do — the markers that hang in the margin
 * and keep the structure readable, and the rule that text typed in the view
 * takes the paragraph style's formatting instead of the run formatting that
 * sits next to it.
 */
import { Editor } from '@tiptap/core'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import type { Node as PmNode } from '@tiptap/pm/model'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  storyHeadingLevel,
  storyIsQuote,
  storyMarkerFor,
  type StoryStyle,
  type StoryStyles,
} from '../src/renderer/editor/story-view'

const STYLES_CSS = readFileSync(join(__dirname, '..', 'src', 'renderer', 'styles.css'), 'utf8')

/** every rule body the stylesheet gives this selector */
function blocksFor(selector: string): string {
  const out: string[] = []
  let at = STYLES_CSS.indexOf(selector)
  while (at >= 0) {
    const open = STYLES_CSS.indexOf('{', at)
    const close = STYLES_CSS.indexOf('}', open)
    out.push(STYLES_CSS.slice(open + 1, close))
    at = STYLES_CSS.indexOf(selector, close)
  }
  return out.join('\n')
}

const makeEditor = (content: unknown): Editor =>
  new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: content as never,
  })

/** the text nodes of the first paragraph, with the marks each one carries */
function textsOf(editor: Editor): Array<{ text: string; marks: string[] }> {
  const out: Array<{ text: string; marks: string[] }> = []
  editor.state.doc.firstChild?.forEach((node: PmNode) => {
    if (!node.isText) return
    out.push({ text: node.text ?? '', marks: node.marks.map((m) => m.type.name) })
  })
  return out
}

/** the end of the first paragraph's text: where a typing cursor would sit */
function endOfFirstParagraph(editor: Editor): number {
  const para = editor.state.doc.firstChild
  return 1 + (para ? para.content.size : 0)
}

/**
 * Type text the way ProseMirror does: handleTextInput runs first (the app's
 * own convention: only real typed text reaches it), and the default insertion
 * follows when no handler took it over.
 */
function typeText(editor: Editor, text: string): void {
  const from = endOfFirstParagraph(editor)
  const to = from
  const deflt = () => editor.state.tr.insertText(text, from, to)
  const handled = editor.view.someProp('handleTextInput', (f) =>
    f(editor.view, from, to, text, deflt),
  )
  if (!handled) editor.view.dispatch(deflt())
}

const HEADING_AND_BODY = {
  type: 'doc',
  content: [
    { type: 'docHeading', attrs: { level: 2 }, content: [{ type: 'text', text: 'Title' }] },
    { type: 'docParagraph', content: [{ type: 'text', text: 'Body' }] },
  ],
}

const STYLED_RUN = {
  type: 'doc',
  content: [
    {
      type: 'docParagraph',
      content: [
        { type: 'text', text: 'Plain ' },
        {
          type: 'text',
          text: 'red',
          marks: [{ type: 'docTextStyle', attrs: { color: 'FF0000' } }],
        },
      ],
    },
  ],
}

describe('story view: the margin markers', () => {
  it('draws one # per heading level and nothing for the body', () => {
    const editor = makeEditor(HEADING_AND_BODY)
    editor.commands.setStoryViewActive(true)
    const markers = [...editor.view.dom.querySelectorAll('.story-marker')].map(
      (el) => el.textContent,
    )
    expect(markers).toEqual(['##'])
    editor.destroy()
  })

  it('draws nothing while the view is off', () => {
    const editor = makeEditor(HEADING_AND_BODY)
    expect(editor.view.dom.querySelectorAll('.story-marker')).toHaveLength(0)
    editor.destroy()
  })

  it('reads the level of a heading that is a styled paragraph', () => {
    const editor = makeEditor({
      type: 'doc',
      content: [
        { type: 'docParagraph', attrs: { styleId: 'H3' }, content: [{ type: 'text', text: 'T' }] },
      ],
    })
    const styles: StoryStyles = new Map([['H3', { name: 'heading 3', headingLevel: 3 }]])
    editor.storage.listNumbering.styles = styles as never
    editor.commands.setStoryViewActive(true)
    const markers = [...editor.view.dom.querySelectorAll('.story-marker')].map(
      (el) => el.textContent,
    )
    expect(markers).toEqual(['###'])
    editor.destroy()
  })
})

describe('story view: the marker rules', () => {
  const node = (name: string, attrs: Record<string, unknown>): PmNode =>
    ({
      type: { name },
      attrs,
    }) as unknown as PmNode

  it('counts a docHeading by its own level', () => {
    expect(storyHeadingLevel(node('docHeading', { level: 3 }))).toBe(3)
    expect(storyMarkerFor(node('docHeading', { level: 3 }))).toBe('###')
  })

  it('counts a paragraph whose style confers a heading level', () => {
    const styles: StoryStyles = new Map([['H2', { name: 'heading 2', headingLevel: 2 }]])
    expect(storyHeadingLevel(node('docParagraph', { styleId: 'H2' }), styles)).toBe(2)
    expect(storyMarkerFor(node('docParagraph', { styleId: 'H2' }), styles)).toBe('##')
  })

  it('counts no level for a style that turns the outline off', () => {
    const styles: StoryStyles = new Map<string, StoryStyle>([
      ['H2', { name: 'heading 2', headingLevel: 2, headingOutlineOff: true }],
    ])
    expect(storyHeadingLevel(node('docParagraph', { styleId: 'H2' }), styles)).toBeNull()
  })

  it('counts no level for a body paragraph', () => {
    expect(storyHeadingLevel(node('docParagraph', { styleId: 'Normal' }))).toBeNull()
    expect(storyMarkerFor(node('docParagraph', { styleId: 'Normal' }))).toBeNull()
  })

  it("marks a quote by Word's built-in quote style", () => {
    const styles: StoryStyles = new Map([
      ['Q', { name: 'Quote' }],
      ['IQ', { name: 'Intense Quote' }],
      ['Q2', { name: 'Citation' }],
    ])
    expect(storyIsQuote(node('docParagraph', { styleId: 'Q' }), styles)).toBe(true)
    expect(storyIsQuote(node('docParagraph', { styleId: 'IQ' }), styles)).toBe(true)
    expect(storyIsQuote(node('docParagraph', { styleId: 'Q2' }), styles)).toBe(false)
    expect(storyMarkerFor(node('docParagraph', { styleId: 'Q' }), styles)).toBe('>')
  })
})

describe('story view: the CSS that hides the formatting', () => {
  // The run formatting rides an inline style and the paragraph formatting an
  // inline style of its own, so the view's rules have to be !important to win.
  it('the run formatting is hidden with !important', () => {
    const body = blocksFor('.view-story .doc-page [data-doc-style]')
    expect(body).toContain('color: inherit !important')
    expect(body).toContain('font-size: inherit !important')
    expect(body).toContain('font-family: inherit !important')
    expect(body).toContain('background-color: transparent !important')
  })

  it('the paragraph indent and spacing are hidden', () => {
    const body = blocksFor('.view-story .doc-page :is(p, h1, h2, h3, h4, h5, h6, .doc-li)')
    expect(body).toContain('margin: 0 !important')
    expect(body).toContain('text-indent: 0 !important')
    expect(body).toContain('padding-inline-start: 0 !important')
  })

  it('the page chrome is gone and the column is centered', () => {
    expect(blocksFor('.view-story .page-hf')).toContain('display: none')
    const page = blocksFor('.view-story .doc-page')
    expect(page).toContain('margin: 0 auto')
    expect(page).toContain('background: none')
    expect(page).toContain('box-shadow: none')
  })

  it('the markers hang in the gutter', () => {
    const body = blocksFor('.view-story .story-marker')
    expect(body).toContain('position: absolute')
    expect(body).toContain('right: 100%')
  })
})

describe('story view: typing takes the paragraph style', () => {
  it('typed text does not pick up the run formatting beside it', () => {
    const editor = makeEditor(STYLED_RUN)
    editor.commands.setStoryViewActive(true)
    typeText(editor, 'q')
    expect(textsOf(editor)).toEqual([
      { text: 'Plain ', marks: [] },
      { text: 'red', marks: ['docTextStyle'] },
      { text: 'q', marks: [] },
    ])
    editor.destroy()
  })

  it('outside the view, typing still inherits the run formatting', () => {
    const editor = makeEditor(STYLED_RUN)
    typeText(editor, 'q')
    expect(textsOf(editor)).toEqual([
      { text: 'Plain ', marks: [] },
      { text: 'redq', marks: ['docTextStyle'] },
    ])
    editor.destroy()
  })

  it('a paste keeps the formatting it carries', () => {
    const editor = makeEditor(STYLED_RUN)
    editor.commands.setStoryViewActive(true)
    // a paste never reaches handleTextInput, so the run formatting stays
    const tr = editor.state.tr.insertText('q', endOfFirstParagraph(editor))
    tr.setMeta('uiEvent', 'paste')
    editor.view.dispatch(tr)
    expect(textsOf(editor)).toEqual([
      { text: 'Plain ', marks: [] },
      { text: 'redq', marks: ['docTextStyle'] },
    ])
    editor.destroy()
  })
})

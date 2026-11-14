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
  storyEscLeavesView,
  storyFocusSentence,
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

  // With Track Changes on, the mark removal on freshly typed text is part of
  // the insertion: an untagged transaction would have the plugin record a
  // formatting edit on top of it, and every typed character would show a
  // "Formatted" bubble in the review pane.
  it('records a typed character as an insertion, not a formatting change', () => {
    const editor = makeEditor(STYLED_RUN)
    editor.storage.trackChanges.enabled = true
    editor.storage.trackChanges.trackFormatting = true
    editor.commands.setStoryViewActive(true)
    typeText(editor, 'q')
    const texts = textsOf(editor)
    expect(texts[texts.length - 1]).toEqual({ text: 'q', marks: ['ins'] })
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

const THREE_SENTENCES = {
  type: 'doc',
  content: [
    {
      type: 'docParagraph',
      content: [{ type: 'text', text: 'First one ends here. Second one goes on. Third one too.' }],
    },
    { type: 'docParagraph', content: [{ type: 'text', text: 'Another block entirely.' }] },
  ],
}

describe('story view Focus: the sentence the caret sits in', () => {
  it('runs from after the previous terminator to the next one', () => {
    const text = 'First one ends here. Second one goes on. Third one too.'
    // the caret on the S of Second: the sentence starts at 21, ends after its dot
    const [from, to] = storyFocusSentence(text, 21)
    // the trailing space of the gap rides the next sentence's range
    expect(text.slice(from, to)).toBe('Second one goes on. ')
  })

  it('a block without a terminator reads as one sentence', () => {
    expect(storyFocusSentence('no stops at all', 5)).toEqual([0, 15])
  })

  it('the CJK terminators end a sentence the same way', () => {
    const text = '\u7B2C\u4E00\u53E5\u3002\u7B2C\u4E8C\uFF01\u7B2C\u4E09'
    const [from, to] = storyFocusSentence(text, 4)
    expect(text.slice(from, to)).toBe('\u7B2C\u4E8C\uFF01')
  })

  it('the caret in the gap after a terminator starts the next sentence', () => {
    const text = 'First one ends here.  Second starts.'
    // the caret on the second space of the gap: the next sentence is active
    const [from, to] = storyFocusSentence(text, 21)
    expect(text.slice(from, to)).toBe('Second starts.')
  })

  it('a dot inside a number does not end a sentence', () => {
    const text = 'Pi is about 3.14 in most books.'
    // the caret after "3.14 in": the dots of 3.14 are not terminators, the
    // block reads as the one sentence the caret sits in
    const [from, to] = storyFocusSentence(text, 17)
    expect(text.slice(from, to)).toBe(text)
  })

  it('a dot inside a host name and a version do not end a sentence', () => {
    const text = 'See example.com for details. Version v2.0.1 ships today.'
    // the caret inside "example": the host's dot is not a terminator, the
    // sentence runs to its own stop
    const [from, to] = storyFocusSentence(text, 7)
    expect(text.slice(from, to)).toBe('See example.com for details. ')
    // the caret inside "v2.0.1": the version's dots keep it whole too
    const [vFrom, vTo] = storyFocusSentence(text, 40)
    expect(text.slice(vFrom, vTo)).toBe('Version v2.0.1 ships today.')
  })

  it('a terminator before a closing quote ends the sentence at the quote', () => {
    const text = 'He said "Stop." Then he left.'
    // the caret on "Then": the previous sentence ends past the closing quote
    const [from, to] = storyFocusSentence(text, 16)
    expect(text.slice(0, from)).toBe('He said "Stop." ')
    expect(text.slice(from, to)).toBe('Then he left.')
  })
})

describe('story view: the Escape that leaves the view', () => {
  const esc = (over: Partial<KeyboardEvent> = {}): KeyboardEvent =>
    ({
      key: 'Escape',
      defaultPrevented: false,
      isComposing: false,
      target: null,
      ...over,
    }) as KeyboardEvent

  it('a plain Escape on the editor leaves the view', () => {
    expect(storyEscLeavesView(esc(), false)).toBe(true)
  })

  it('an Escape of an open IME composition does not', () => {
    expect(storyEscLeavesView(esc({ isComposing: true }), false)).toBe(false)
  })

  it('an Escape a transient overlay already consumed does not', () => {
    expect(storyEscLeavesView(esc({ defaultPrevented: true }), false)).toBe(false)
  })

  it('an Escape while an overlay that owns it is open does not', () => {
    // the context menu, the comment popup, the dialogs each close their own Escape
    expect(storyEscLeavesView(esc(), true)).toBe(false)
  })

  it('an Escape inside the find panel belongs to the panel', () => {
    const host = document.createElement('div')
    host.className = 'find-panel'
    document.body.append(host)
    expect(storyEscLeavesView(esc({ target: host }), false)).toBe(false)
    host.remove()
  })

  it('another key never leaves the view', () => {
    expect(storyEscLeavesView(esc({ key: 'Enter' }), false)).toBe(false)
  })
})

describe('story view Focus: the dim ranges', () => {
  /** the text the view dims, with the caret parked inside the second sentence */
  function dimmedText(editor: Editor): string {
    return [...editor.view.dom.querySelectorAll('.story-dim')]
      .map((el) => el.textContent)
      .join('\n')
  }

  it("dims every sentence but the caret's, in every block", () => {
    const editor = makeEditor(THREE_SENTENCES)
    editor.commands.setStoryViewActive(true)
    editor.commands.setStoryViewFocus(true)
    editor.commands.setTextSelection(23) // inside "Second"
    const dim = dimmedText(editor)
    expect(dim).toContain('First one ends here.')
    expect(dim).toContain('Third one too.')
    expect(dim).toContain('Another block entirely.')
    expect(dim).not.toContain('Second')
    editor.destroy()
  })

  it('follows the caret: moving repaints the dim around the new sentence', () => {
    const editor = makeEditor(THREE_SENTENCES)
    editor.commands.setStoryViewActive(true)
    editor.commands.setStoryViewFocus(true)
    editor.commands.setTextSelection(2) // inside "First"
    expect(dimmedText(editor)).not.toContain('First')
    expect(dimmedText(editor)).toContain('Second one goes on.')
    editor.destroy()
  })

  it('dims nothing while Focus is off, and nothing outside the view', () => {
    const editor = makeEditor(THREE_SENTENCES)
    editor.commands.setStoryViewActive(true)
    editor.commands.setStoryViewFocus(false)
    expect(editor.view.dom.querySelectorAll('.story-dim')).toHaveLength(0)
    editor.commands.setStoryViewFocus(true)
    editor.commands.setStoryViewActive(false)
    expect(editor.view.dom.querySelectorAll('.story-dim')).toHaveLength(0)
    editor.destroy()
  })

  it('a block holding an inline object stays visible when it is the active one', () => {
    const editor = makeEditor({
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [{ type: 'text', text: 'Plain words here. More words.' }],
        },
        {
          type: 'docParagraph',
          content: [{ type: 'text', text: 'With an object. ' }, { type: 'docNoteRef' }],
        },
      ],
    } as never)
    editor.commands.setStoryViewActive(true)
    editor.commands.setStoryViewFocus(true)
    editor.commands.setTextSelection(3) // inside the first block
    // the second block (inline object) dims whole; the active block's sentence
    // logic ran without touching positions it could not compute
    expect(dimmedText(editor)).toContain('With an object.')
    editor.destroy()
  })
})

describe('story view frame: the chrome leaves and Focus dims', () => {
  it('the ribbon and the panes are hidden in the story frame', () => {
    expect(blocksFor('.story-chrome .ribbon')).toContain('display: none')
    expect(blocksFor('.story-chrome .ai-dock')).toContain('display: none')
    expect(blocksFor('.story-chrome .nav-pane')).toContain('display: none')
    expect(blocksFor('.story-chrome .comments-pane')).toContain('display: none')
  })

  it('the find panel stays available in the story frame', () => {
    // Ctrl/Cmd+F still sets showFind here, so the frame hides the chrome
    // panes but not the panel — hiding it made the shortcut read as dead
    expect(blocksFor('.story-chrome .find-panel')).toBe('')
  })

  it('the story canvas uses the themed canvas, not the print-preview one', () => {
    expect(blocksFor('.story-chrome .editor-scroll')).toContain('var(--canvas)')
  })

  it('the crop marks belong to the page and do not show here', () => {
    expect(blocksFor('.view-story .page-wrap::before')).toContain('display: none')
    expect(blocksFor('.view-story .page-gap-paper::after')).toContain('display: none')
  })

  it('the dim rides a decoration class scoped to the view', () => {
    expect(blocksFor('.view-story .story-dim')).toContain('opacity: 0.25')
  })

  it('the face comes from the writing fonts when they are installed', () => {
    expect(blocksFor('.view-story')).toContain('Monaspace Neon')
    expect(blocksFor('.view-story')).toContain('IBM Plex Mono')
    expect(blocksFor('.view-story')).toContain('system-ui')
  })

  it('the bar is plain text labels in one row', () => {
    expect(blocksFor('.story-bar')).toContain('position: fixed')
    expect(blocksFor('.story-bar-btn')).toContain('font: inherit')
  })
})

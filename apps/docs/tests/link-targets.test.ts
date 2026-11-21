/**
 * Insert Hyperlink beyond a web address: the places of this document a link can
 * point at, the hidden bookmark a link to a heading rides on, and e-mail links.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { generateParagraphXml } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { pmNodeToGeneratedBlock } from '../src/renderer/editor/convert'
import { linkTarget } from '../src/renderer/editor/link-actions'
import {
  anchorPos,
  buildMailto,
  collectLinkTargets,
  ensureHeadingAnchor,
  headingAnchorName,
  linkKindOf,
  parseMailto,
} from '../src/renderer/editor/link-targets'

const live: Editor[] = []
afterEach(() => {
  for (const e of live.splice(0)) e.destroy()
})

const text = (t: string) => [{ type: 'text', text: t }]

function makeEditor() {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: text('intro') },
        { type: 'docHeading', attrs: { level: 1 }, content: text('Scope & Goals') },
        {
          type: 'docHeading',
          attrs: { level: 2, hiddenBookmarks: ['_GoBack', '_Toc123456789'] },
          content: text('Details'),
        },
        { type: 'docParagraph', attrs: { bookmarks: ['figures'] }, content: text('see figures') },
        { type: 'docHeading', attrs: { level: 1 }, content: text('Scope & Goals') },
      ],
    },
  })
  live.push(editor)
  return editor
}

describe('collectLinkTargets', () => {
  it('lists headings with their anchors and the user bookmarks', () => {
    const { headings, bookmarks } = collectLinkTargets(makeEditor().state.doc)
    expect(headings.map((h) => [h.level, h.text, h.anchors])).toEqual([
      [1, 'Scope & Goals', []],
      [2, 'Details', ['_GoBack', '_Toc123456789']],
      [1, 'Scope & Goals', []],
    ])
    expect(bookmarks.map((b) => [b.name, b.preview])).toEqual([['figures', 'see figures']])
  })
})

describe('heading anchors', () => {
  it('names a bookmark after the heading text', () => {
    expect(headingAnchorName('Scope & Goals', new Set())).toBe('_Scope_Goals')
    expect(headingAnchorName('  1.2  Scope ', new Set())).toBe('_1_2_Scope')
    expect(headingAnchorName('!!!', new Set())).toBe('_Heading')
    expect(headingAnchorName('Scope', new Set(['_scope', '_scope_1']))).toBe('_Scope_2')
    expect(headingAnchorName('x'.repeat(80), new Set()).length).toBeLessThanOrEqual(40)
  })

  it('keeps letters of any script', () => {
    expect(headingAnchorName('\u7b2c\u4e00\u7ae0 \u6982\u8ff0', new Set())).toBe(
      '_\u7b2c\u4e00\u7ae0_\u6982\u8ff0',
    )
  })

  it('adds a hidden bookmark to a heading that has none, once', () => {
    const editor = makeEditor()
    const first = collectLinkTargets(editor.state.doc).headings[0].pos
    const tr = editor.state.tr
    expect(ensureHeadingAnchor(tr, first)).toBe('_Scope_Goals')
    editor.view.dispatch(tr)
    expect(editor.state.doc.nodeAt(first)?.attrs.hiddenBookmarks).toEqual(['_Scope_Goals'])
    const again = editor.state.tr
    expect(ensureHeadingAnchor(again, first)).toBe('_Scope_Goals')
    expect(again.docChanged).toBe(false)
    expect(anchorPos(editor.state.doc, '_Scope_Goals')).toBe(first)
  })

  it('numbers the bookmark of a second heading with the same text', () => {
    const editor = makeEditor()
    const { headings } = collectLinkTargets(editor.state.doc)
    const tr = editor.state.tr
    expect(ensureHeadingAnchor(tr, headings[0].pos)).toBe('_Scope_Goals')
    expect(ensureHeadingAnchor(tr, headings[2].pos)).toBe('_Scope_Goals_1')
    editor.view.dispatch(tr)
    const again = editor.state.tr
    expect(ensureHeadingAnchor(again, headings[2].pos)).toBe('_Scope_Goals_1')
    expect(again.docChanged).toBe(false)
  })

  it('reuses a table-of-contents bookmark, not an unrelated hidden one', () => {
    const editor = makeEditor()
    const details = collectLinkTargets(editor.state.doc).headings[1].pos
    const tr = editor.state.tr
    expect(ensureHeadingAnchor(tr, details)).toBe('_Toc123456789')
    expect(tr.docChanged).toBe(false)
  })

  it('keeps the anchor of the link being edited when the heading carries it', () => {
    const editor = makeEditor()
    const details = collectLinkTargets(editor.state.doc).headings[1].pos
    const tr = editor.state.tr
    expect(ensureHeadingAnchor(tr, details, '_GoBack')).toBe('_GoBack')
    expect(ensureHeadingAnchor(tr, details, 'elsewhere')).toBe('_Toc123456789')
    expect(tr.docChanged).toBe(false)
  })

  it('refuses a position that is not a heading', () => {
    const editor = makeEditor()
    expect(ensureHeadingAnchor(editor.state.tr, 0)).toBeNull()
  })

  it('the bookmark and the internal link are written to the paragraph XML', () => {
    const editor = makeEditor()
    const first = collectLinkTargets(editor.state.doc).headings[0].pos
    const tr = editor.state.tr
    const name = ensureHeadingAnchor(tr, first)
    editor.view.dispatch(tr)
    editor
      .chain()
      .setTextSelection(2)
      .insertContent({
        type: 'text',
        text: 'go',
        marks: [{ type: 'link', attrs: { href: `#${name}`, rId: null, tooltip: 'Jump' } }],
      })
      .run()
    const ctx = { headingStyleIds: new Map<number, string>(), allocateHyperlinkRel: () => 'rId9' }
    const xml = (i: number) =>
      generateParagraphXml(pmNodeToGeneratedBlock(editor.state.doc.child(i).toJSON()), ctx)
    expect(xml(0)).toContain('<w:hyperlink w:anchor="_Scope_Goals" w:tooltip="Jump">')
    expect(xml(1)).toMatch(/<w:bookmarkStart [^>]*w:name="_Scope_Goals"/)
  })
})

describe('anchorPos', () => {
  it('finds user and hidden bookmarks, null for an unknown name', () => {
    const doc = makeEditor().state.doc
    expect(doc.nodeAt(anchorPos(doc, 'figures')!)?.textContent).toBe('see figures')
    expect(doc.nodeAt(anchorPos(doc, '_Toc123456789')!)?.textContent).toBe('Details')
    expect(anchorPos(doc, 'nope')).toBeNull()
  })
})

describe('e-mail links', () => {
  it('builds and reads back a mailto address with its subject', () => {
    const href = buildMailto({ address: ' team@example.com ', subject: 'Q3 report & plan' })
    expect(href).toBe('mailto:team@example.com?subject=Q3%20report%20%26%20plan')
    expect(parseMailto(href)).toEqual({ address: 'team@example.com', subject: 'Q3 report & plan' })
  })

  it('leaves the subject out when empty and accepts a pasted mailto: prefix', () => {
    expect(buildMailto({ address: 'mailto:a@example.com', subject: ' ' })).toBe(
      'mailto:a@example.com',
    )
    expect(buildMailto({ address: '  ', subject: 'x' })).toBe('')
  })

  it('reads an address written by another program', () => {
    expect(parseMailto('MAILTO:a@example.com?cc=b@example.com&Subject=Hi%20there')).toEqual({
      address: 'a@example.com',
      subject: 'Hi there',
    })
    expect(parseMailto('mailto:a@example.com?subject=100%')).toEqual({
      address: 'a@example.com',
      subject: '100%',
    })
    expect(parseMailto('https://example.com')).toBeNull()
  })

  it('opens in the mail program, and only for a real address', () => {
    expect(linkTarget('mailto:a@example.com?subject=Hi')).toEqual({
      kind: 'external',
      url: 'mailto:a@example.com?subject=Hi',
    })
    expect(linkTarget('mailto:')).toBeNull()
    expect(linkTarget('mailto:nobody')).toBeNull()
  })
})

describe('linkKindOf', () => {
  it('picks the tab an existing link opens on', () => {
    expect(linkKindOf('#_Toc1')).toBe('doc')
    expect(linkKindOf('mailto:a@example.com')).toBe('email')
    expect(linkKindOf('https://example.com')).toBe('web')
    expect(linkKindOf('')).toBe('web')
  })
})

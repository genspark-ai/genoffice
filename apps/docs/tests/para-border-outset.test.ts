import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { paraBorderOutset, paraBorderOutsetDecls } from '../src/renderer/editor/para-border-css'
import { borderGroupKey } from '../src/renderer/editor/para-border-merge'

function styleOf(type: string, attrs: Record<string, unknown>): string[] {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type, attrs: { docxIndex: null, ...attrs }, content: [{ type: 'text', text: 'x' }] },
      ],
    },
  })
  const style = editor.view.dom.firstElementChild!.getAttribute('style') ?? ''
  editor.destroy()
  return style.split(';').map((d) => d.trim().replace(/:\s+/, ':'))
}

const box = (l: number, r: number) => ({
  borders: 'tblr',
  borderLines: JSON.stringify({
    t: {},
    b: {},
    l: l ? { spacePt: l } : {},
    r: r ? { spacePt: r } : {},
  }),
})

describe('left / right w:space of a drawn paragraph border', () => {
  it('reads the gap of drawn sides only', () => {
    expect(paraBorderOutset('lr', { l: { spacePt: 4 }, r: { spacePt: 6 } })).toEqual({ l: 4, r: 6 })
    expect(paraBorderOutset('l', { l: { spacePt: 4 }, r: { spacePt: 6 } })).toEqual({ l: 4, r: 0 })
    expect(paraBorderOutset('tb', { t: { spacePt: 4 }, b: { spacePt: 4 } })).toBeNull()
    expect(paraBorderOutset('lr', { l: {}, r: {} })).toBeNull()
    expect(paraBorderOutset('', { l: { none: true, spacePt: 31 } })).toBeNull()
  })

  it('moves the box out and pads it back, so the text column keeps its width', () => {
    const style = styleOf('docParagraph', box(4, 6))
    expect(style).toContain('margin-inline-start:calc(var(--pind-s, 0pt) - 4pt)')
    expect(style).toContain('padding-inline-start:4pt')
    expect(style).toContain('margin-inline-end:calc(var(--pind-e, 0pt) - 6pt)')
    expect(style).toContain('padding-inline-end:6pt')
  })

  it('leaves a border without a side gap as it was', () => {
    const style = styleOf('docParagraph', box(0, 0)).join(';')
    expect(style).not.toContain('margin-inline')
    expect(style).not.toContain('padding-inline')
  })

  it('moves the paragraph own indent', () => {
    const style = styleOf('docParagraph', { ...box(4, 6), indentLeft: 720, indentRight: 360 })
    expect(style).toContain('margin-inline-start:32pt')
    expect(style).toContain('margin-inline-end:12pt')
    expect(style.indexOf('margin-inline-start:32pt')).toBeGreaterThan(
      style.indexOf('margin-inline-start:36pt'),
    )
  })

  it('keeps the left line on the left of a right-to-left paragraph', () => {
    const style = styleOf('docParagraph', { ...box(4, 0), bidi: true })
    expect(style).toContain('padding-inline-end:4pt')
    expect(style.join(';')).not.toContain('padding-inline-start')
  })

  it('widens a list item through its indent padding', () => {
    const style = styleOf('docListItem', { kind: 'bullet', ilvl: 0, ...box(4, 6) })
    expect(style).toContain('--pbdr-os:4pt')
    expect(style).toContain('margin-inline-start:calc(var(--pind-s, 0pt) - 4pt)')
    expect(style.join(';')).not.toContain('padding-inline-start')
    expect(style).toContain('padding-inline-end:6pt')
  })

  it('hands a hosted anchor the moved margin', () => {
    const own = ['margin-inline-start:36pt', '--anchor-ind:36pt']
    expect(paraBorderOutsetDecls(own, { l: 4, r: 0 }, { anchorHost: true })).toContain(
      '--anchor-ind:32pt',
    )
    // right to left: the left line is the end edge, the right one leaves the left edge alone
    const rtl = { anchorHost: true, rtl: true }
    expect(paraBorderOutsetDecls(own, { l: 4, r: 0 }, rtl)).toContain('--anchor-ind:32pt')
    expect(paraBorderOutsetDecls(own, { l: 0, r: 6 }, rtl).join(';')).not.toContain('--anchor-ind')
  })

  it('groups only paragraphs with the same side gap', () => {
    const key = (l: number) => borderGroupKey(box(l, 0))
    expect(key(4)).toBe(key(4))
    expect(key(4)).not.toBe(key(5))
  })
})

import { describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import { buildBlankDocx, parseDocx, type StyleInfo } from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  activeStyleKey,
  allStyleEntries,
  collectUsedStyleIds,
  headingStyleId,
  quickStyleEntries,
  styleLabel,
  stylePreviewCss,
} from '../src/renderer/style-gallery'

const WORD_GALLERY = [
  'Normal',
  'No Spacing',
  'heading 1',
  'heading 2',
  'Title',
  'Subtitle',
  'Subtle Emphasis',
  'Emphasis',
  'Intense Emphasis',
  'Strong',
  'Quote',
  'Intense Quote',
  'Subtle Reference',
  'Intense Reference',
  'Book Title',
  'List Paragraph',
]

async function blankStyles() {
  return (await parseDocx(await buildBlankDocx())).styles
}

const style = (over: Partial<StyleInfo> & { styleId: string }): StyleInfo => ({
  name: over.styleId,
  type: 'paragraph',
  ...over,
})
const mapOf = (...infos: StyleInfo[]) => new Map(infos.map((s) => [s.styleId, s]))

describe('quick style gallery (Word order and hiding)', () => {
  it('a new document lists Word’s 16 quick styles in uiPriority order', async () => {
    const styles = await blankStyles()
    expect(quickStyleEntries(styles, new Set()).map((s) => s.name)).toEqual(WORD_GALLERY)
  })

  it('Heading 3 joins right after Heading 2 once the document uses it', async () => {
    const styles = await blankStyles()
    const names = quickStyleEntries(styles, new Set(['Heading3'])).map((s) => s.name)
    expect(names.slice(2, 5)).toEqual(['heading 1', 'heading 2', 'heading 3'])
    expect(names).toHaveLength(17)
  })

  it('All styles lists the hidden headings and TOC levels too, never linked character shells', async () => {
    const styles = await blankStyles()
    const all = allStyleEntries(styles).map((s) => s.styleId)
    expect(all).toContain('Heading6')
    expect(all).toContain('TOC1')
    expect(all).toContain('Hyperlink')
  })

  it('orders by uiPriority, then by name (numeric aware), default style always present', () => {
    const styles = mapOf(
      style({ styleId: 'Z', name: 'Zeta', qFormat: true, uiPriority: 5 }),
      style({ styleId: 'H10', name: 'heading 10', qFormat: true, uiPriority: 9 }),
      style({ styleId: 'H9', name: 'heading 9', qFormat: true, uiPriority: 9 }),
      style({ styleId: 'Body', name: 'Body', isDefault: true }),
      style({ styleId: 'Hidden', name: 'Hidden', qFormat: true, semiHidden: true }),
      style({
        styleId: 'Shell',
        name: 'Shell',
        type: 'character',
        qFormat: true,
        linkedCharShell: true,
      }),
      style({ styleId: 'Tbl', name: 'Grid', type: 'table', qFormat: true }),
    )
    expect(quickStyleEntries(styles, new Set(['Hidden'])).map((s) => s.styleId)).toEqual([
      'Body',
      'Z',
      'H9',
      'H10',
    ])
  })

  it('semiHidden styles unhide only with unhideWhenUsed and a use in the document', () => {
    const styles = mapOf(
      style({ styleId: 'Normal', isDefault: true, qFormat: true }),
      style({ styleId: 'A', qFormat: true, semiHidden: true, unhideWhenUsed: true, uiPriority: 1 }),
      style({ styleId: 'B', qFormat: true, semiHidden: true, uiPriority: 2 }),
    )
    expect(quickStyleEntries(styles, new Set()).map((s) => s.styleId)).toEqual(['Normal'])
    expect(quickStyleEntries(styles, new Set(['A', 'B'])).map((s) => s.styleId)).toEqual([
      'Normal',
      'A',
    ])
  })

  it('falls back to the visible styles when nothing carries qFormat', () => {
    const styles = mapOf(
      style({ styleId: 'Normal', isDefault: true }),
      style({ styleId: 'Body2', name: 'Body Text' }),
      style({ styleId: 'Secret', semiHidden: true }),
    )
    expect(quickStyleEntries(styles, new Set()).map((s) => s.styleId)).toEqual(['Body2', 'Normal'])
  })

  it('finds the document’s heading style by level even under a localized id', () => {
    const styles = mapOf(
      style({ styleId: '1', name: 'heading 1', headingLevel: 1 }),
      style({ styleId: 'Custom', name: 'Chapter', headingLevel: 2 }),
      style({ styleId: 'Child', name: 'Child', headingLevel: 2, headingLevelInherited: true }),
    )
    expect(headingStyleId(styles, 1)).toBe('1')
    expect(headingStyleId(styles, 2)).toBe('Custom')
    expect(headingStyleId(styles, 3)).toBeUndefined()
  })

  it('collects used hidden styles from headings, paragraph styles and runs', async () => {
    const styles = new Map(await blankStyles())
    styles.set(
      'HiddenChar',
      style({
        styleId: 'HiddenChar',
        type: 'character',
        qFormat: true,
        semiHidden: true,
        unhideWhenUsed: true,
      }),
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docHeading', attrs: { level: 3 }, content: [{ type: 'text', text: 'h3' }] },
          {
            type: 'docParagraph',
            attrs: { styleId: 'Heading4' },
            content: [{ type: 'text', text: 'p' }],
          },
          {
            type: 'docParagraph',
            content: [
              {
                type: 'text',
                text: 'run',
                marks: [{ type: 'docTextStyle', attrs: { styleId: 'HiddenChar' } }],
              },
            ],
          },
        ],
      },
    })
    const used = collectUsedStyleIds(editor.state.doc, styles)
    expect([...used].sort()).toEqual(['Heading3', 'Heading4', 'HiddenChar'])
    expect(collectUsedStyleIds(null, styles).size).toBe(0)
    editor.destroy()
  })

  it('highlights the character style at the caret, else the paragraph style', async () => {
    const styles = await blankStyles()
    expect(
      activeStyleKey({ charStyleId: 'Strong', paraStyleId: 'Title', headingLevel: null }, styles),
    ).toBe('char:Strong')
    expect(
      activeStyleKey({ charStyleId: null, paraStyleId: 'Title', headingLevel: null }, styles),
    ).toBe('para:Title')
    expect(activeStyleKey({ charStyleId: null, paraStyleId: null, headingLevel: 2 }, styles)).toBe(
      'para:Heading2',
    )
    expect(
      activeStyleKey({ charStyleId: null, paraStyleId: null, headingLevel: null }, styles),
    ).toBe('para:Normal')
  })

  it('labels built-ins like Word and previews the resolved formatting', async () => {
    const styles = await blankStyles()
    const t = (key: string, params?: Record<string, string | number>) =>
      key === 'ribbonStyleHeadingN' ? `H ${params?.n}` : key
    expect(styleLabel(styles.get('Heading1')!, t as never)).toBe('ribbonStyleHeading1')
    expect(styleLabel({ name: 'heading 7' }, t as never)).toBe('H 7')
    expect(styleLabel({ name: 'caption' }, t as never)).toBe('Caption')
    const intense = stylePreviewCss(styles.get('IntenseEmphasis')!)
    expect(intense.fontStyle).toBe('italic')
    expect(intense.color).toBe('#2F5496')
    expect(stylePreviewCss(styles.get('IntenseReference')!).fontVariant).toBe('small-caps')
    expect(stylePreviewCss(styles.get('Title')!).fontSize).toBe('20px')
  })
})

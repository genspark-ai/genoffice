/**
 * Font dialog: underline style and colour, emphasis mark, OpenType features and
 * Default…. What the dialog writes reaches the run model, and what it leaves
 * alone stays as each run had it.
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { describe, expect, it, vi } from 'vitest'
import type { Run } from '@genoffice/docx-engine'
import {
  FontDialog,
  selectionRanges,
  setTextStyleWithScale,
  setUnderlineMark,
} from '../src/renderer/components/FontDialog'
import { inlineToRuns, runsToInline } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import {
  emphasisDecls,
  otfFeatureSettings,
  underlineDecls,
} from '../src/renderer/editor/run-typography'
import { LocaleProvider } from '../src/renderer/i18n/locale'

Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })

function editorWith(runs: Run[]): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [{ type: 'docParagraph', attrs: { docxIndex: 0 }, content: runsToInline(runs) }],
    },
  })
}

const runsOf = (editor: Editor): Run[] =>
  inlineToRuns(editor.getJSON().content![0].content as Parameters<typeof inlineToRuns>[0])

function selectAll(editor: Editor) {
  const { doc } = editor.state
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(doc, 1, doc.content.size - 1)),
  )
}

function openDialog(editor: Editor, onDefault = vi.fn()) {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() =>
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(FontDialog, { editor, onDefault, onClose: () => {} }),
      }),
    ),
  )
  const pick = (label: string, value: string) => {
    const trigger = container.querySelector<HTMLButtonElement>(`.gs-dd-btn[aria-label="${label}"]`)!
    act(() => trigger.click())
    act(() =>
      container.querySelector<HTMLButtonElement>(`.gs-dd-item[data-value="${value}"]`)!.click(),
    )
  }
  const click = (text: string) =>
    act(() =>
      [...container.querySelectorAll('button')].find((b) => b.textContent === text)!.click(),
    )
  const close = () => {
    act(() => root.unmount())
    container.remove()
  }
  return { container, pick, click, close, onDefault }
}

describe('CSS for the run typography', () => {
  it('draws each underline style; a plain single line needs nothing', () => {
    expect(underlineDecls(null, null)).toEqual([])
    expect(underlineDecls('single', 'auto')).toEqual([])
    expect(underlineDecls('double', null)).toEqual(['text-decoration-style:double'])
    expect(underlineDecls('thick', null)).toEqual(['text-decoration-thickness:0.12em'])
    expect(underlineDecls('dotted', null)).toEqual(['text-decoration-style:dotted'])
    expect(underlineDecls('dottedHeavy', null)).toEqual([
      'text-decoration-style:dotted',
      'text-decoration-thickness:0.12em',
    ])
    expect(underlineDecls('dash', null)).toEqual(['text-decoration-style:dashed'])
    expect(underlineDecls('dotDotDash', null)).toEqual(['text-decoration-style:dashed'])
    expect(underlineDecls('wave', 'FF0000')).toEqual([
      'text-decoration-style:wavy',
      'text-decoration-color:#FF0000',
    ])
    expect(underlineDecls('wavyDouble', null)).toEqual(['text-decoration-style:wavy'])
  })

  it('turns OpenType features into font-feature-settings', () => {
    expect(otfFeatureSettings(undefined)).toBeNull()
    expect(otfFeatureSettings({})).toBeNull()
    expect(otfFeatureSettings({ ligatures: 'standardContextual' })).toBe(
      '"liga" 1,"clig" 1,"hlig" 0,"dlig" 0',
    )
    expect(otfFeatureSettings({ ligatures: 'none' })).toBe('"liga" 0,"clig" 0,"hlig" 0,"dlig" 0')
    expect(otfFeatureSettings({ ligatures: 'all' })).toBe('"liga" 1,"clig" 1,"hlig" 1,"dlig" 1')
    expect(
      otfFeatureSettings({
        numForm: 'oldStyle',
        numSpacing: 'tabular',
        stylisticSets: [3],
        cntxtAlts: true,
      }),
    ).toBe('"onum" 1,"tnum" 1,"ss03" 1,"calt" 1')
  })

  it('draws no emphasis mark for none', () => {
    expect(emphasisDecls('none')).toEqual([])
    expect(emphasisDecls('dot')).toEqual([
      'text-emphasis:filled dot',
      'text-emphasis-position:under right',
    ])
  })
})

describe('run <-> mark round trip', () => {
  it('carries the underline style and colour, the emphasis mark and the OpenType features', () => {
    const otf = { ligatures: 'standard', numForm: 'lining' as const, stylisticSets: [2] }
    const editor = editorWith([
      { text: 'wavy', underline: true, underlineStyle: 'wave', underlineColor: 'C00000' },
      { text: ' plain ', underline: true },
      { text: 'marked', em: 'dot', otf },
    ])
    const dom = editor.view.dom
    expect(dom.querySelector('u')!.getAttribute('style')).toBe(
      'text-decoration-style: wavy; text-decoration-color: rgb(192, 0, 0);',
    )
    // a plain line renders as it always did
    expect(dom.querySelectorAll('u')[1].attributes).toHaveLength(0)
    const style = dom.querySelector('span[data-doc-style]')!.getAttribute('style')
    expect(style).toContain('"liga" 1,"clig" 0,"hlig" 0,"dlig" 0,"lnum" 1,"ss02" 1')
    expect(style).toContain('text-emphasis: filled dot')

    const [wavy, plain, marked] = runsOf(editor)
    expect(wavy).toMatchObject({
      underline: true,
      underlineStyle: 'wave',
      underlineColor: 'C00000',
    })
    // the editor's run spells a plain line out
    expect(plain).toMatchObject({
      underline: true,
      underlineStyle: 'single',
      underlineColor: 'auto',
    })
    expect(marked).toMatchObject({ em: 'dot', otf })
    expect(marked.underline).toBeUndefined()
    editor.destroy()
  })
})

describe('setUnderlineMark', () => {
  it('a colour alone recolours underlined text and underlines nothing else', () => {
    const editor = editorWith([
      { text: 'double', underline: true, underlineStyle: 'double' },
      { text: ' bare' },
    ])
    selectAll(editor)
    editor
      .chain()
      .command(({ tr }) =>
        setUnderlineMark(tr, selectionRanges(editor.state.selection), { color: '0070C0' }),
      )
      .run()
    const [double, bare] = runsOf(editor)
    expect(double).toMatchObject({ underlineStyle: 'double', underlineColor: '0070C0' })
    expect(bare.underline).toBeUndefined()
    editor.destroy()
  })

  it('a style goes on every run and keeps each colour; none removes the line', () => {
    const editor = editorWith([
      { text: 'red', underline: true, underlineColor: 'FF0000' },
      { text: ' bare' },
    ])
    selectAll(editor)
    const ranges = selectionRanges(editor.state.selection)
    editor
      .chain()
      .command(({ tr }) => setUnderlineMark(tr, ranges, { line: 'dotted' }))
      .run()
    expect(runsOf(editor)).toMatchObject([
      { underlineStyle: 'dotted', underlineColor: 'FF0000' },
      { underlineStyle: 'dotted', underlineColor: 'auto' },
    ])
    editor
      .chain()
      .command(({ tr }) => setUnderlineMark(tr, ranges, { line: 'none' }))
      .run()
    expect(runsOf(editor).every((r) => r.underline === undefined)).toBe(true)
    editor.destroy()
  })
})

describe('OpenType features over a mixed selection', () => {
  it('sets one feature and leaves the others each run had', () => {
    const editor = editorWith([{ text: 'old', otf: { numForm: 'oldStyle' } }, { text: ' bare' }])
    selectAll(editor)
    editor
      .chain()
      .command(({ tr }) =>
        setTextStyleWithScale(
          tr,
          selectionRanges(editor.state.selection),
          {},
          {
            ligatures: 'all',
          },
        ),
      )
      .run()
    expect(runsOf(editor)).toMatchObject([
      { otf: { ligatures: 'all', numForm: 'oldStyle' } },
      { otf: { ligatures: 'all' } },
    ])
    editor.destroy()
  })
})

describe('the dialog', () => {
  it('applies underline style, emphasis mark and typography, and only what was touched', () => {
    const editor = editorWith([
      { text: 'kept', em: 'comma', otf: { numSpacing: 'tabular' } },
      { text: ' bare' },
    ])
    selectAll(editor)
    const dialog = openDialog(editor)
    dialog.pick('Underline style', 'double')
    dialog.click('Advanced')
    dialog.pick('Ligatures', 'standardContextual')
    dialog.pick('Stylistic sets', '4')
    dialog.click('OK')
    expect(runsOf(editor)).toMatchObject([
      {
        underlineStyle: 'double',
        em: 'comma',
        otf: { ligatures: 'standardContextual', numSpacing: 'tabular', stylisticSets: [4] },
      },
      { underlineStyle: 'double', otf: { ligatures: 'standardContextual', stylisticSets: [4] } },
    ])
    expect(runsOf(editor)[1].em).toBeUndefined()
    dialog.close()
    editor.destroy()
  })

  it('shows the selection, removes the emphasis mark and does not offer Snap to grid', () => {
    const editor = editorWith([
      { text: 'marked', em: 'dot', underline: true, underlineStyle: 'wavyDouble' },
    ])
    selectAll(editor)
    const dialog = openDialog(editor)
    const trigger = (label: string) =>
      dialog.container.querySelector<HTMLButtonElement>(`.gs-dd-btn[aria-label="${label}"]`)!
    expect(trigger('Emphasis mark').textContent).toContain('Dot')
    // a line the list does not offer is still the current value
    act(() => trigger('Underline style').click())
    expect(dialog.container.querySelector('.gs-dd-item[data-value="wavyDouble"]')).not.toBeNull()
    act(() => trigger('Underline style').click())
    dialog.pick('Emphasis mark', 'none')
    dialog.click('Advanced')
    const snap = [...dialog.container.querySelectorAll('label')].find((l) =>
      l.textContent?.startsWith('Snap to grid'),
    )
    expect(snap).toBeUndefined()
    dialog.click('OK')
    expect(runsOf(editor)[0]).toMatchObject({ em: 'none', underlineStyle: 'wavyDouble' })
    dialog.close()
    editor.destroy()
  })

  it('Default… hands the font over and applies the dialog', () => {
    const editor = editorWith([{ text: 'body', fontAscii: 'Georgia', sizeHalfPoints: 28 }])
    selectAll(editor)
    const dialog = openDialog(editor)
    dialog.pick('Font style', 'boldItalic')
    dialog.click('Default…')
    expect(dialog.onDefault).toHaveBeenCalledWith({
      latin: 'Georgia',
      sizeHalfPoints: 28,
      bold: true,
      italic: true,
    })
    expect(editor.isActive('bold')).toBe(true)
    dialog.close()
    editor.destroy()
  })
})

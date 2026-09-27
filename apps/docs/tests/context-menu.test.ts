import { describe, expect, it, vi } from 'vitest'
import { createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { act } from 'react'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { LocaleProvider, setModuleLang } from '../src/renderer/i18n/locale'
import {
  EditorContextMenu,
  FontDialog,
  ParagraphDialog,
} from '../src/renderer/components/ContextMenu'

function createEditor(paraAttrs: Record<string, unknown> = {}): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0, ...paraAttrs },
          content: [{ type: 'text', text: 'EVs market research' }],
        },
      ],
    },
  })
}

function select(editor: Editor, from: number, to: number) {
  editor.view.dispatch(
    editor.state.tr.setSelection(TextSelection.create(editor.state.doc, from, to)),
  )
}

/** Drive a shared Dropdown (gs-dd): open the trigger, click the option by value. */
function pickDropdown(container: Element, dd: HTMLButtonElement, value: string) {
  act(() => dd.click())
  const item = container.querySelector<HTMLButtonElement>(`.gs-dd-item[data-value="${value}"]`)!
  act(() => item.click())
}

function render(element: React.ReactElement): { container: HTMLElement; unmount: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(createElement(LocaleProvider, { initial: 'en', children: element })))
  return {
    container,
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

const noop = () => {}

function menuProps(editor: Editor, overrides: Record<string, unknown> = {}) {
  return {
    editor,
    menu: { x: 10, y: 10 },
    onClose: noop,
    onFontDialog: noop,
    onParagraphDialog: noop,
    onLink: noop,
    onNewComment: noop,
    onViewImage: noop,
    onSaveImageAs: noop,
    onAiPreset: noop,
    ...overrides,
  }
}

// LocaleProvider subscribes to the shell's language switch on mount
Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })

setModuleLang('en')

describe('EditorContextMenu', () => {
  it('disables selection-dependent items when nothing is selected', () => {
    const editor = createEditor()
    const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
    const byLabel = (label: string) =>
      [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
        (b) => b.querySelector('.ctx-label')?.textContent === label,
      )!
    expect(byLabel('Cut').disabled).toBe(true)
    expect(byLabel('Copy').disabled).toBe(true)
    expect(byLabel('Paste').disabled).toBe(false)
    expect(byLabel('Font…').disabled).toBe(false)
    expect(byLabel('Paragraph…').disabled).toBe(false)
    // Word anchors a comment on the word under a collapsed caret, so this stays live
    expect(byLabel('New Comment').disabled).toBe(false)
    unmount()
    editor.destroy()
  })

  it('disables New Comment when the caret has no word to anchor on', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', attrs: { docxIndex: 0 } }],
      },
    })
    const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
    const item = [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
      (b) => b.querySelector('.ctx-label')?.textContent === 'New Comment',
    )!
    expect(item.disabled).toBe(true)
    unmount()
    editor.destroy()
  })

  it('enables everything and routes New Comment / Font… when text is selected', () => {
    const editor = createEditor()
    select(editor, 1, 5)
    const onNewComment = vi.fn()
    const onFontDialog = vi.fn()
    const onClose = vi.fn()
    const { container, unmount } = render(
      createElement(EditorContextMenu, menuProps(editor, { onNewComment, onFontDialog, onClose })),
    )
    const byLabel = (label: string) =>
      [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
        (b) => b.querySelector('.ctx-label')?.textContent === label,
      )!
    expect(byLabel('Cut').disabled).toBe(false)
    expect(byLabel('New Comment').disabled).toBe(false)
    act(() => byLabel('New Comment').click())
    expect(onNewComment).toHaveBeenCalledOnce()
    expect(onClose).toHaveBeenCalled()
    act(() => byLabel('Font…').click())
    expect(onFontDialog).toHaveBeenCalledOnce()
    unmount()
    editor.destroy()
  })

  it('sends the selected text to the AI panel for Synonyms', () => {
    const editor = createEditor()
    select(editor, 1, 4)
    const onAiPreset = vi.fn()
    const { container, unmount } = render(
      createElement(EditorContextMenu, menuProps(editor, { onAiPreset })),
    )
    const synonym = [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
      (b) => b.querySelector('.ctx-label')?.textContent === 'Synonyms',
    )!
    expect(synonym.disabled).toBe(false)
    act(() => synonym.click())
    expect(onAiPreset).toHaveBeenCalledOnce()
    expect(String(onAiPreset.mock.calls[0][0])).toContain('EVs')
    unmount()
    editor.destroy()
  })

  it('marks AI-backed items (Synonyms/Translate) with the copilot badge', () => {
    const editor = createEditor()
    select(editor, 1, 4)
    const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
    const badged = [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')]
      .filter((b) => b.querySelector('.copilot-badge'))
      .map((b) => b.querySelector('.ctx-label')?.textContent)
    expect(badged).toContain('Synonyms')
    expect(badged).toContain('Translate')
    unmount()
    editor.destroy()
  })
})

describe('FontDialog', () => {
  it('applies font marks to the selection on OK', () => {
    const editor = createEditor()
    select(editor, 1, 10)
    const { container, unmount } = render(createElement(FontDialog, { editor, onClose: noop }))
    const dds = container.querySelectorAll<HTMLButtonElement>('.gs-dd-btn')
    // Latin font, East Asian font, font style → bold, size → 12
    pickDropdown(container, dds[0]!, 'Arial')
    pickDropdown(container, dds[1]!, '\u5b8b\u4f53')
    pickDropdown(container, dds[2]!, 'bold')
    pickDropdown(container, dds[3]!, '12')
    const ok = [...container.querySelectorAll('button')].find((b) => b.textContent === 'OK')!
    act(() => ok.click())
    expect(editor.isActive('bold')).toBe(true)
    const attrs = editor.getAttributes('docTextStyle')
    expect(attrs.sizeHalfPoints).toBe(24)
    // untouched fields stay as they were
    expect(attrs.color).toBeNull()
    // each picker writes only its own rFonts slot
    expect(attrs.fontAscii).toBe('Arial')
    expect(attrs.font).toBe('\u5b8b\u4f53')
    expect(attrs.eastAsiaFont).toBe('\u5b8b\u4f53')
    unmount()
    editor.destroy()
  })
})

describe('ParagraphDialog', () => {
  it('applies alignment and spacing to the paragraph on OK', () => {
    const editor = createEditor()
    select(editor, 2, 2)
    const { container, unmount } = render(createElement(ParagraphDialog, { editor, onClose: noop }))
    pickDropdown(container, container.querySelector<HTMLButtonElement>('.gs-dd-btn')!, 'center')
    const ok = [...container.querySelectorAll('button')].find((b) => b.textContent === 'OK')!
    act(() => ok.click())
    expect(editor.getAttributes('docParagraph').align).toBe('center')
    unmount()
    editor.destroy()
  })

  it('resolves visual left/right against the paragraph direction (LTR)', () => {
    const editor = createEditor({ align: 'right' })
    select(editor, 2, 2)
    const { container, unmount } = render(createElement(ParagraphDialog, { editor, onClose: noop }))
    const alignDd = container.querySelector<HTMLButtonElement>('.gs-dd-btn')!
    expect(alignDd.dataset.value).toBe('right')
    pickDropdown(container, alignDd, 'left')
    const ok = [...container.querySelectorAll('button')].find((b) => b.textContent === 'OK')!
    act(() => ok.click())
    // visual left is the start side in LTR → stored as null
    expect(editor.getAttributes('docParagraph').align).toBeNull()
    unmount()
    editor.destroy()
  })

  it('shows the start side as Right in RTL and stores visual left explicitly', () => {
    const editor = createEditor({ bidi: true })
    select(editor, 2, 2)
    const { container, unmount } = render(createElement(ParagraphDialog, { editor, onClose: noop }))
    const alignDd = container.querySelector<HTMLButtonElement>('.gs-dd-btn')!
    // unset align in an RTL paragraph renders right, so the dialog shows Right
    expect(alignDd.dataset.value).toBe('right')
    pickDropdown(container, alignDd, 'left')
    const ok = [...container.querySelectorAll('button')].find((b) => b.textContent === 'OK')!
    act(() => ok.click())
    // visual left is the end side in RTL → stored explicitly
    expect(editor.getAttributes('docParagraph').align).toBe('left')
    unmount()
    editor.destroy()
  })

  it('clears the align attr when re-selecting the start side in RTL', () => {
    const editor = createEditor({ bidi: true, align: 'left' })
    select(editor, 2, 2)
    const { container, unmount } = render(createElement(ParagraphDialog, { editor, onClose: noop }))
    const alignDd = container.querySelector<HTMLButtonElement>('.gs-dd-btn')!
    expect(alignDd.dataset.value).toBe('left')
    pickDropdown(container, alignDd, 'right')
    const ok = [...container.querySelectorAll('button')].find((b) => b.textContent === 'OK')!
    act(() => ok.click())
    expect(editor.getAttributes('docParagraph').align).toBeNull()
    unmount()
    editor.destroy()
  })
})

describe('EditorContextMenu picture items', () => {
  const labels = (container: HTMLElement) =>
    [...container.querySelectorAll('.ctx-label')].map((el) => el.textContent)

  it('shows View / Save Image As only when the click landed on a picture', () => {
    const editor = createEditor()
    const plain = render(createElement(EditorContextMenu, menuProps(editor)))
    expect(labels(plain.container)).not.toContain('View Image')
    plain.unmount()

    const onViewImage = vi.fn()
    const onSaveImageAs = vi.fn()
    const src = 'data:image/png;base64,AAAA'
    const { container, unmount } = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, { menu: { x: 10, y: 10, imageSrc: src }, onViewImage, onSaveImageAs }),
      ),
    )
    const names = labels(container)
    expect(names.slice(0, 2)).toEqual(['View Image', 'Save Image As…'])
    const byLabel = (label: string) =>
      [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
        (b) => b.querySelector('.ctx-label')?.textContent === label,
      )!
    byLabel('View Image').click()
    byLabel('Save Image As…').click()
    expect(onViewImage).toHaveBeenCalledWith(src)
    expect(onSaveImageAs).toHaveBeenCalledWith(src)
    unmount()
  })
})

describe('EditorContextMenu — Word hyperlink / field / spelling items', () => {
  const byLabel = (container: Element, label: string) =>
    [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
      (b) => b.querySelector('.ctx-label')?.textContent === label,
    )
  const labels = (container: Element) =>
    [...container.querySelectorAll('.ctx-label')].map((el) => el.textContent)

  function linkEditor() {
    return new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [
              { type: 'text', text: 'see ' },
              {
                type: 'text',
                text: 'the site',
                marks: [{ type: 'link', attrs: { href: 'https://example.com/a' } }],
              },
              { type: 'text', text: ' now' },
            ],
          },
        ],
      },
    })
  }

  it('replaces Hyperlink… with Edit / Open / Copy / Remove on a link run', () => {
    const editor = linkEditor()
    const onOpenLink = vi.fn()
    const onLink = vi.fn()
    const writeText = vi.fn(() => Promise.resolve())
    Object.assign(navigator, { clipboard: { writeText } })
    // "see " = 1..5, link = 5..13; the pointer position is what counts, not the caret
    const { container, unmount } = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, { menu: { x: 10, y: 10, pos: 8 }, onOpenLink, onLink }),
      ),
    )
    const names = labels(container)
    expect(names).not.toContain('Hyperlink…')
    expect(names.slice(-5)).toEqual([
      'Edit Hyperlink…',
      'Open Hyperlink',
      'Copy Hyperlink',
      'Remove Hyperlink',
      'New Comment',
    ])
    act(() => byLabel(container, 'Open Hyperlink')!.click())
    expect(onOpenLink).toHaveBeenCalledWith('https://example.com/a')
    act(() => byLabel(container, 'Copy Hyperlink')!.click())
    expect(writeText).toHaveBeenCalledWith('https://example.com/a')
    act(() => byLabel(container, 'Edit Hyperlink…')!.click())
    expect(onLink).toHaveBeenCalledOnce()
    expect([editor.state.selection.from, editor.state.selection.to]).toEqual([5, 13])
    act(() => byLabel(container, 'Remove Hyperlink')!.click())
    expect(editor.state.doc.textContent).toBe('see the site now')
    expect(editor.state.doc.nodeAt(6)!.marks).toHaveLength(0)
    unmount()
    editor.destroy()
  })

  it('offers only Open / Copy on a TOC entry and keeps Hyperlink… off links', () => {
    const editor = linkEditor()
    const toc = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, {
          menu: { x: 10, y: 10, pos: 2, link: { href: '#_Toc1', toc: true, tocTitle: 'Intro' } },
          onOpenLink: () => {},
        }),
      ),
    )
    expect(byLabel(toc.container, 'Edit Hyperlink…')!.disabled).toBe(true)
    expect(byLabel(toc.container, 'Open Hyperlink')!.disabled).toBe(false)
    expect(byLabel(toc.container, 'Remove Hyperlink')!.disabled).toBe(true)
    toc.unmount()

    const plain = render(
      createElement(EditorContextMenu, menuProps(editor, { menu: { x: 10, y: 10, pos: 2 } })),
    )
    expect(labels(plain.container)).toContain('Hyperlink…')
    expect(labels(plain.container)).not.toContain('Open Hyperlink')
    plain.unmount()
    editor.destroy()
  })

  it('adds Toggle Field Codes and Edit Field… on an inline field', () => {
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [
              { type: 'text', text: 'Page ' },
              {
                type: 'text',
                text: '7',
                marks: [{ type: 'instrField', attrs: { instr: 'PAGE' } }],
              },
            ],
          },
        ],
      },
    })
    const onEditField = vi.fn()
    const onUpdateFields = vi.fn()
    const { container, unmount } = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, { menu: { x: 10, y: 10, pos: 6 }, onEditField, onUpdateFields }),
      ),
    )
    const names = labels(container)
    const at = names.indexOf('Update Field')
    expect(names.slice(at, at + 3)).toEqual(['Update Field', 'Toggle Field Codes', 'Edit Field…'])
    act(() => byLabel(container, 'Edit Field…')!.click())
    expect(onEditField).toHaveBeenCalledWith(expect.objectContaining({ instr: 'PAGE', from: 6 }))
    act(() => byLabel(container, 'Toggle Field Codes')!.click())
    expect(editor.view.dom.querySelector('.doc-field-code')?.textContent).toBe('{ PAGE }')
    unmount()
    editor.destroy()
  })

  it('lists Ignore All before Add to Dictionary and routes it to the session skip list', async () => {
    const editor = createEditor()
    const spellIgnoreWord = vi.fn(() => Promise.resolve(true))
    const spellAddWord = vi.fn(() => Promise.resolve(true))
    Object.assign(window.desktop, { spellIgnoreWord, spellAddWord })
    const onRespell = vi.fn()
    const { container, unmount } = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, {
          menu: { x: 10, y: 10, pos: 2, spell: { word: 'EVs', suggestions: ['Eve'] } },
          onRespell,
        }),
      ),
    )
    const names = labels(container)
    expect(names.slice(0, 3)).toEqual(['Eve', 'Ignore All', 'Add to Dictionary'])
    act(() => byLabel(container, 'Ignore All')!.click())
    expect(spellIgnoreWord).toHaveBeenCalledWith('EVs')
    expect(spellAddWord).not.toHaveBeenCalled()
    await act(async () => {})
    expect(onRespell).toHaveBeenCalled()
    unmount()
    editor.destroy()
  })
})

import { describe, expect, it, vi } from 'vitest'
import { createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { act } from 'react'
import { Editor } from '@tiptap/core'
import { NodeSelection, TextSelection } from '@tiptap/pm/state'
import { CellSelection } from '@tiptap/pm/tables'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { setGrammarFindings, setGrammarMarksHidden } from '../src/renderer/editor/grammar-marks'
import { LocaleProvider, setModuleLang } from '../src/renderer/i18n/locale'
import {
  EditorContextMenu,
  FontDialog,
  ParagraphDialog,
} from '../src/renderer/components/ContextMenu'
import { insertTableAt } from '../src/renderer/components/ribbon-tabs'
import { ignoreSpellingOnce } from '../src/renderer/editor/spell-menu'

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
    onCommand: noop,
    onViewImage: noop,
    onSaveImageAs: noop,
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

  it('Synonyms and Translate act on the word a bare caret is in, and are off beside no word', () => {
    const editor = createEditor()
    editor.commands.setTextSelection(7)
    const onCommand = vi.fn()
    const { container, unmount } = render(
      createElement(EditorContextMenu, menuProps(editor, { onCommand })),
    )
    const byLabel = (label: string) =>
      [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
        (b) => b.querySelector('.ctx-label')?.textContent === label,
      )!
    expect(byLabel('Synonyms').disabled).toBe(false)
    expect(byLabel('Translate…').disabled).toBe(false)
    // one item as in Word: the language is the one Tools > Translate is set to
    act(() => byLabel('Translate…').click())
    expect(onCommand.mock.calls).toEqual([['tools-translate-selection']])
    unmount()
    editor.destroy()

    const spaced = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          {
            type: 'docParagraph',
            attrs: { docxIndex: 0 },
            content: [{ type: 'text', text: 'one   two' }],
          },
        ],
      },
    })
    spaced.commands.setTextSelection(6)
    const off = render(createElement(EditorContextMenu, menuProps(spaced)))
    const offLabel = (label: string) =>
      [...off.container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
        (b) => b.querySelector('.ctx-label')?.textContent === label,
      )!
    expect(offLabel('Synonyms').disabled).toBe(true)
    expect(offLabel('Translate…').disabled).toBe(true)
    off.unmount()
    spaced.destroy()
  })

  it('New Comment is live in an empty paragraph and off at a caret beside no word', () => {
    const newComment = (paragraph: Record<string, unknown>) => {
      const editor = new Editor({
        element: document.createElement('div'),
        extensions: editorExtensions,
        content: { type: 'doc', content: [{ type: 'docParagraph', ...paragraph }] },
      })
      const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
      const item = [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
        (b) => b.querySelector('.ctx-label')?.textContent === 'New Comment',
      )!
      const { disabled } = item
      unmount()
      editor.destroy()
      return disabled
    }
    expect(newComment({ attrs: { docxIndex: 0 } })).toBe(false)
    expect(newComment({ content: [{ type: 'text', text: '... ...' }] })).toBe(true)
  })

  it('enables everything and routes New Comment / Font… when text is selected', () => {
    const editor = createEditor()
    select(editor, 1, 5)
    const onCommand = vi.fn()
    const onClose = vi.fn()
    const { container, unmount } = render(
      createElement(EditorContextMenu, menuProps(editor, { onCommand, onClose })),
    )
    const byLabel = (label: string) =>
      [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
        (b) => b.querySelector('.ctx-label')?.textContent === label,
      )!
    expect(byLabel('Cut').disabled).toBe(false)
    expect(byLabel('New Comment').disabled).toBe(false)
    act(() => byLabel('New Comment').click())
    expect(onCommand).toHaveBeenLastCalledWith('insert-comment')
    expect(onClose).toHaveBeenCalled()
    act(() => byLabel('Font…').click())
    expect(onCommand).toHaveBeenLastCalledWith('font-dialog')
    unmount()
    editor.destroy()
  })

  it('on a grammar mark: the message, the suggestions, Ignore Once and Grammar\u2026, all through commands', () => {
    const editor = createEditor()
    setGrammarFindings(editor, [
      {
        id: 'g-ctx',
        from: 5,
        to: 11,
        quote: 'market',
        message: 'Use the plural here.',
        suggestions: ['markets', 'market\u2019s'],
        key: 'k',
      },
    ])
    const onCommand = vi.fn()
    const labelsAt = (pos: number, mount = editor) => {
      const { container, unmount } = render(
        createElement(
          EditorContextMenu,
          menuProps(mount, { onCommand, menu: { x: 10, y: 10, pos } }),
        ),
      )
      const items = [...container.querySelectorAll<HTMLButtonElement>('.ctx-menu > .ctx-item')]
      return { items, labels: items.map((b) => b.textContent), unmount }
    }
    const on = labelsAt(7)
    expect(on.labels.slice(0, 5)).toEqual([
      'Use the plural here.',
      'markets',
      'market\u2019s',
      'Ignore Once',
      'Grammar\u2026',
    ])
    expect(on.items[0].disabled).toBe(true)
    act(() => on.items[1].click())
    expect(onCommand).toHaveBeenLastCalledWith(
      'grammar-replace',
      JSON.stringify({ id: 'g-ctx', text: 'markets' }),
    )
    act(() => on.items[3].click())
    expect(onCommand).toHaveBeenLastCalledWith('grammar-ignore', 'g-ctx')
    act(() => on.items[4].click())
    expect(onCommand).toHaveBeenLastCalledWith('grammar-dialog', 'g-ctx')
    on.unmount()

    // off the mark, and while the marks are hidden, the menu is the plain one
    const off = labelsAt(2)
    expect(off.labels).not.toContain('Ignore Once')
    off.unmount()
    setGrammarMarksHidden(editor, true)
    const hidden = labelsAt(7)
    expect(hidden.labels).not.toContain('Ignore Once')
    hidden.unmount()
    setGrammarMarksHidden(editor, false)

    // a document that cannot be edited still shows the finding; its suggestions are greyed out
    editor.setEditable(false)
    const locked = labelsAt(7)
    expect(locked.labels[1]).toBe('markets')
    expect(locked.items[1].disabled).toBe(true)
    expect(locked.items[3].disabled).toBe(false)
    locked.unmount()
    editor.destroy()
  })

  it('lists synonym candidates in a submenu, asked for when it opens, and replaces through a command', async () => {
    const editor = createEditor()
    select(editor, 1, 4)
    const onCommand = vi.fn()
    const aiChat = vi.fn(async (_request: unknown) => ({
      ok: true,
      content: '1. electric cars\n- EVs\n"e-vehicles"',
    }))
    Object.assign(window.desktop, { getAiSettings: async () => ({}), aiChat })
    const { container, unmount } = render(
      createElement(EditorContextMenu, menuProps(editor, { onCommand })),
    )
    const labels = () =>
      [...container.querySelectorAll('.ctx-submenu .ctx-label')].map((el) => el.textContent)
    const synonym = [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
      (b) => b.querySelector('.ctx-label')?.textContent === 'Synonyms',
    )!
    expect(synonym.disabled).toBe(false)
    expect(aiChat).not.toHaveBeenCalled()
    await act(async () => {
      synonym.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }))
    })
    expect(aiChat).toHaveBeenCalledOnce()
    expect(aiChat.mock.calls[0][0]).toMatchObject({ user: 'EVs' })
    expect(labels()).toEqual(['electric cars', 'e-vehicles', 'Thesaurus\u2026'])
    const pick = [...container.querySelectorAll<HTMLButtonElement>('.ctx-submenu .ctx-item')]
    act(() => pick[0].click())
    expect(onCommand).toHaveBeenLastCalledWith('synonym-replace', 'electric cars')
    unmount()
    editor.destroy()
  })

  it('a read-only document shows the candidates greyed and keeps the Thesaurus live', async () => {
    const editor = createEditor()
    select(editor, 1, 4)
    editor.setEditable(false)
    Object.assign(window.desktop, {
      getAiSettings: async () => ({}),
      aiChat: async () => ({ ok: true, content: 'electric cars' }),
    })
    const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
    const synonym = [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
      (b) => b.querySelector('.ctx-label')?.textContent === 'Synonyms',
    )!
    await act(async () => {
      synonym.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }))
    })
    const items = [...container.querySelectorAll<HTMLButtonElement>('.ctx-submenu .ctx-item')]
    expect(items.map((b) => [b.textContent, b.disabled])).toEqual([
      ['electric cars', true],
      ['Thesaurus\u2026', false],
    ])
    unmount()
    editor.destroy()
  })

  it('says No Suggestions when the request fails and still offers the Thesaurus', async () => {
    const editor = createEditor()
    select(editor, 1, 4)
    const onCommand = vi.fn()
    Object.assign(window.desktop, {
      getAiSettings: async () => ({}),
      aiChat: async () => ({ ok: false, error: 'no key' }),
    })
    const { container, unmount } = render(
      createElement(EditorContextMenu, menuProps(editor, { onCommand })),
    )
    const synonym = [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
      (b) => b.querySelector('.ctx-label')?.textContent === 'Synonyms',
    )!
    await act(async () => {
      synonym.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }))
    })
    const items = [...container.querySelectorAll<HTMLButtonElement>('.ctx-submenu .ctx-item')]
    expect(items.map((b) => b.textContent)).toEqual(['(No Suggestions)', 'Thesaurus\u2026'])
    expect(items[0].disabled).toBe(true)
    act(() => items[1].click())
    expect(onCommand).toHaveBeenLastCalledWith('tools-thesaurus')
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
    expect(badged).toContain('Translate…')
    unmount()
    editor.destroy()
  })
})

describe('FontDialog', () => {
  it('applies font marks to the selection on OK', () => {
    const editor = createEditor()
    select(editor, 1, 10)
    const { container, unmount } = render(
      createElement(FontDialog, { editor, onDefault: noop, onClose: noop }),
    )
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

/** the top-level rows of the open menu, a separator as '---' */
const outline = (container: HTMLElement) =>
  [...container.querySelector('.ctx-menu')!.children].map((row) =>
    row.classList.contains('ctx-sep')
      ? '---'
      : (row.matches('.ctx-item') ? row : row.querySelector(':scope > .ctx-item'))?.querySelector(
          '.ctx-label',
        )?.textContent,
  )

describe('EditorContextMenu picture items', () => {
  const labels = (container: HTMLElement) =>
    [...container.querySelectorAll('.ctx-label')].map((el) => el.textContent)

  it('a picture gets the picture menu Word shows, in the same order', () => {
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
    const byLabel = (label: string) =>
      [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
        (b) => b.querySelector('.ctx-label')?.textContent === label,
      )!
    const names = labels(container)
    // Word's picture menu: the clipboard trio first, then the picture items,
    // and none of the text-only entries
    expect(names).toEqual([
      'Cut',
      'Copy',
      'Paste',
      'Save as Picture…',
      'Change Picture',
      'Group',
      'Bring to Front',
      'Send to Back',
      'Hyperlink…',
      'Insert Caption…',
      'Wrap Text',
      'View Alt Text…',
      'Crop',
      'Auto Crop',
      'Size and Position…',
      'Format Picture…',
      'View Image',
    ])
    // the items that act on the selected picture are greyed out while only the pointer is on one
    for (const later of [
      'Change Picture',
      'Group',
      'View Alt Text…',
      'Auto Crop',
      'Size and Position…',
      'Format Picture…',
    ]) {
      expect(byLabel(later).disabled, later).toBe(true)
    }
    byLabel('View Image').click()
    byLabel('Save as Picture…').click()
    expect(onViewImage).toHaveBeenCalledWith(src)
    expect(onSaveImageAs).toHaveBeenCalledWith(src)
    // Word's groups: a line after Change Picture and one on each side of Hyperlink…
    expect(outline(container)).toEqual([
      'Cut',
      'Copy',
      'Paste',
      '---',
      'Save as Picture…',
      'Change Picture',
      '---',
      'Group',
      'Bring to Front',
      'Send to Back',
      '---',
      'Hyperlink…',
      '---',
      'Insert Caption…',
      'Wrap Text',
      '---',
      'View Alt Text…',
      'Crop',
      'Auto Crop',
      'Size and Position…',
      'Format Picture…',
      '---',
      'View Image',
    ])
    unmount()
  })

  it('Change Picture offers the file, the icons and the clipboard; the clipboard only when it holds a picture', async () => {
    const src = 'data:image/png;base64,AAAA'
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [
          { type: 'docProtected', attrs: { blockType: 'image', imageDataUrl: src } },
          { type: 'docParagraph', content: [{ type: 'text', text: 'body' }] },
        ],
      },
    })
    editor.view.dispatch(editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, 0)))
    const open = async (types: string[]) => {
      Object.assign(navigator, { clipboard: { read: () => Promise.resolve([{ types }]) } })
      const onCommand = vi.fn()
      const shown = render(
        createElement(
          EditorContextMenu,
          menuProps(editor, { menu: { x: 10, y: 10, imageSrc: src }, onCommand }),
        ),
      )
      await act(async () => {})
      const change = [...shown.container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
        (b) => b.querySelector('.ctx-label')?.textContent === 'Change Picture',
      )!
      act(() => {
        change.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }))
      })
      const items = [
        ...shown.container.querySelectorAll<HTMLButtonElement>('.ctx-submenu .ctx-item'),
      ]
      return { ...shown, items, onCommand }
    }

    const text = await open(['text/plain'])
    expect(text.items.map((b) => b.textContent)).toEqual([
      'From a File…',
      'From Icons…',
      'From Clipboard…',
    ])
    expect(text.items.map((b) => b.disabled)).toEqual([false, false, true])
    act(() => text.items[1].click())
    expect(text.onCommand).toHaveBeenLastCalledWith('picture-replace-icon')
    text.unmount()

    // a picture in a format the save cannot write as it is does not count
    const other = await open(['image/webp'])
    expect(other.items[2].disabled).toBe(true)
    other.unmount()

    const picture = await open(['text/html', 'image/png'])
    expect(picture.items[2].disabled).toBe(false)
    act(() => picture.items[2].click())
    expect(picture.onCommand).toHaveBeenLastCalledWith('picture-replace-clipboard')
    picture.unmount()
    editor.destroy()
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
    const onCommand = vi.fn()
    const writeText = vi.fn(() => Promise.resolve())
    Object.assign(navigator, { clipboard: { writeText } })
    // "see " = 1..5, link = 5..13; the pointer position is what counts, not the caret
    const { container, unmount } = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, { menu: { x: 10, y: 10, pos: 8 }, onOpenLink, onCommand }),
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
    expect(onCommand).toHaveBeenLastCalledWith('insert-link')
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

  it('a marked word lists its items in Word order and runs each through a command', () => {
    const editor = createEditor()
    const onCommand = vi.fn()
    const { container, unmount } = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, {
          menu: { x: 10, y: 10, pos: 2, spell: { word: 'EVs', suggestions: ['Eve'] } },
          onCommand,
        }),
      ),
    )
    expect(labels(container).slice(0, 6)).toEqual([
      'Eve',
      'Ignore Once',
      'Ignore All',
      'Add to Dictionary',
      'AutoCorrect',
      'Spelling…',
    ])
    act(() => byLabel(container, 'Ignore All')!.click())
    expect(onCommand).toHaveBeenLastCalledWith('spell-ignore-all', 'EVs')
    act(() => byLabel(container, 'Add to Dictionary')!.click())
    expect(onCommand).toHaveBeenLastCalledWith('spell-add-word', 'EVs')
    act(() => byLabel(container, 'Eve')!.click())
    expect(onCommand.mock.lastCall![0]).toBe('spell-replace')
    expect(JSON.parse(onCommand.mock.lastCall![1])).toEqual({
      pos: 2,
      word: 'EVs',
      to: 'Eve',
      marked: true,
    })
    unmount()
    editor.destroy()
  })

  it('an occurrence set aside by Ignore Once has no spelling section, even while the browser still marks it', () => {
    const editor = createEditor()
    ignoreSpellingOnce(editor, { from: 1, to: 4 })
    const { container, unmount } = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, {
          menu: { x: 10, y: 10, pos: 2, spell: { word: 'EVs', suggestions: ['Eve'] } },
        }),
      ),
    )
    expect(labels(container)[0]).toBe('Cut')
    unmount()
    editor.destroy()
  })

  it('a misspelled word the browser never marked gets the same items from the dictionary', () => {
    const editor = createEditor()
    Object.assign(window.desktop, {
      spellIsMisspelled: (word: string) => word === 'EVs',
      spellSuggestions: () => ['Eves', 'Evs'],
    })
    const onCommand = vi.fn()
    // a bare caret inside the word, no data from the browser
    editor.commands.setTextSelection(2)
    const { container, unmount } = render(
      createElement(EditorContextMenu, menuProps(editor, { onCommand })),
    )
    expect(labels(container).slice(0, 3)).toEqual(['Eves', 'Evs', 'Ignore Once'])
    act(() => byLabel(container, 'Spelling…')!.click())
    expect(onCommand).toHaveBeenLastCalledWith('tools-spelling', '1')
    // the AutoCorrect submenu repeats the suggestions
    const auto = byLabel(container, 'AutoCorrect')!
    act(() => {
      auto.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }))
    })
    const pair = [...container.querySelectorAll<HTMLButtonElement>('.ctx-submenu .ctx-item')]
    expect(pair.map((b) => b.textContent)).toEqual(['Eves', 'Evs'])
    act(() => pair[1].click())
    expect(onCommand.mock.lastCall![0]).toBe('spell-autocorrect')
    expect(JSON.parse(onCommand.mock.lastCall![1])).toMatchObject({ word: 'EVs', to: 'Evs' })
    unmount()

    // a word that is spelled right gets no spelling section
    editor.commands.setTextSelection(8)
    const plain = render(createElement(EditorContextMenu, menuProps(editor)))
    expect(labels(plain.container)[0]).toBe('Cut')
    plain.unmount()
    Object.assign(window.desktop, { spellIsMisspelled: undefined, spellSuggestions: undefined })
    editor.destroy()
  })
})

describe('EditorContextMenu — comment, note and tracked-change scenes', () => {
  const labels = (container: Element) =>
    [...container.querySelectorAll('.ctx-label')].map((el) => el.textContent)
  const byLabel = (container: Element, label: string) =>
    [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
      (b) => b.querySelector('.ctx-label')?.textContent === label,
    )

  it('a comment balloon shows Reply, Delete and Resolve only; a resolved one Reopen', () => {
    const editor = createEditor()
    const onCommand = vi.fn()
    const open = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, {
          menu: { x: 10, y: 10, pos: 2, commentId: 'c1', balloon: true },
          comment: { id: 'c1', done: false },
          onCommand,
        }),
      ),
    )
    expect(labels(open.container)).toEqual([
      'Reply To Comment',
      'Delete Comment',
      'Resolve Comment',
    ])
    act(() => byLabel(open.container, 'Resolve Comment')!.click())
    expect(onCommand).toHaveBeenLastCalledWith('comment-resolve', 'c1')
    open.unmount()

    const done = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, {
          menu: { x: 10, y: 10, pos: 2, commentId: 'c1', balloon: true },
          comment: { id: 'c1', done: true },
          onCommand,
        }),
      ),
    )
    expect(labels(done.container)).toEqual(['Reply To Comment', 'Delete Comment', 'Reopen Comment'])
    expect(byLabel(done.container, 'Reply To Comment')!.disabled).toBe(true)
    act(() => byLabel(done.container, 'Reopen Comment')!.click())
    expect(onCommand).toHaveBeenLastCalledWith('comment-reopen', 'c1')
    done.unmount()
    editor.destroy()
  })

  it('the anchor text keeps the text menu with Reply, Delete and Resolve live, as Word lists them', () => {
    const editor = createEditor()
    const onCommand = vi.fn()
    const { container, unmount } = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, {
          menu: { x: 10, y: 10, pos: 2 },
          comment: { id: 'c1', done: false },
          onCommand,
        }),
      ),
    )
    expect(outline(container)).toEqual([
      'Cut',
      'Copy',
      'Paste',
      'Paste Special…',
      '---',
      'Reply To Comment',
      'Delete Comment',
      'Resolve Comment',
      '---',
      'Font…',
      'Paragraph…',
      '---',
      'Synonyms',
      'Translate…',
      '---',
      'Hyperlink…',
      'New Comment',
    ])
    act(() => byLabel(container, 'Reply To Comment')!.click())
    expect(onCommand).toHaveBeenLastCalledWith('comment-reply', 'c1')
    unmount()
    editor.destroy()
  })

  it('the withholding item keeps its place after Translate\u2026, and the document-wide items Word lacks are gone', () => {
    const editor = createEditor()
    select(editor, 1, 4)
    const { container, unmount } = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, { comment: { id: null, done: false }, onRedact: noop }),
      ),
    )
    const rows = outline(container)
    expect(rows.slice(rows.indexOf('Synonyms'))).toEqual([
      'Synonyms',
      'Translate…',
      'Hide the selection from AI',
      '---',
      'Hyperlink…',
      'New Comment',
    ])
    for (const gone of ['Paste as Plain Text', 'Edit comment', 'Track Changes'])
      expect(rows, gone).not.toContain(gone)
    unmount()
    editor.destroy()
  })

  it('a note text gets the note items alone; its mark in the body adds them to the text menu', () => {
    const editor = createEditor()
    const onCommand = vi.fn()
    const text = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, {
          menu: { x: 10, y: 10, note: { kind: 'footnote', id: '4', at: 'text' } },
          onCommand,
        }),
      ),
    )
    expect(labels(text.container)).toEqual([
      'Go to Footnote',
      'Note Options…',
      'Convert to Endnote',
    ])
    act(() => byLabel(text.container, 'Go to Footnote')!.click())
    expect(onCommand).toHaveBeenLastCalledWith('note-goto', 'footnote:ref:4')
    text.unmount()

    const mark = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, {
          menu: { x: 10, y: 10, pos: 2, note: { kind: 'endnote', id: '2', at: 'ref' } },
          onCommand,
        }),
      ),
    )
    const names = labels(mark.container)
    expect(names[0]).toBe('Cut')
    expect(names).toEqual(
      expect.arrayContaining(['Go to Endnote', 'Note Options…', 'Convert to Footnote']),
    )
    act(() => byLabel(mark.container, 'Convert to Footnote')!.click())
    expect(onCommand).toHaveBeenLastCalledWith('note-convert', 'endnote:2')
    act(() => byLabel(mark.container, 'Note Options…')!.click())
    expect(onCommand).toHaveBeenLastCalledWith('insert-footnote-dialog', undefined)
    mark.unmount()
    editor.destroy()
  })

  it('a tracked insertion or deletion is named and offered at the top', () => {
    const editor = createEditor()
    const onRange = vi.fn()
    const revisionsOf = (kind: 'ins' | 'del' | 'rPrChange') => ({
      at: [{ from: 1, to: 4, kind }],
      any: true,
      onRange,
      onAll: noop,
    })
    const ins = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, { menu: { x: 10, y: 10, pos: 2 }, revisions: revisionsOf('ins') }),
      ),
    )
    expect(labels(ins.container).slice(0, 3)).toEqual([
      'Accept Insertion',
      'Reject Insertion',
      'Cut',
    ])
    act(() => byLabel(ins.container, 'Accept Insertion')!.click())
    expect(onRange).toHaveBeenLastCalledWith('accept', expect.objectContaining({ kind: 'ins' }))
    // the document-wide items stay where they were, at the end
    expect(labels(ins.container).slice(-2)).toEqual(['Accept All Changes', 'Reject All Changes'])
    ins.unmount()

    const del = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, { menu: { x: 10, y: 10, pos: 2 }, revisions: revisionsOf('del') }),
      ),
    )
    expect(labels(del.container).slice(0, 2)).toEqual(['Accept Deletion', 'Reject Deletion'])
    del.unmount()

    const fmt = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, { menu: { x: 10, y: 10, pos: 2 }, revisions: revisionsOf('rPrChange') }),
      ),
    )
    expect(labels(fmt.container).slice(0, 2)).toEqual(['Accept Change', 'Reject Change'])
    fmt.unmount()
    editor.destroy()
  })
})

describe('EditorContextMenu \u2014 table items run the shared table commands', () => {
  const tableEditor = () => {
    const editor = createEditor()
    insertTableAt(editor, 2, 2)
    const cells: number[] = []
    editor.state.doc.descendants((node, pos) => {
      if (node.type.spec.tableRole?.endsWith('cell')) cells.push(pos)
    })
    return { editor, cells }
  }
  const selectCells = (editor: Editor, anchor: number, head: number) =>
    editor.view.dispatch(
      editor.state.tr.setSelection(CellSelection.create(editor.state.doc, anchor, head)),
    )
  const open = (editor: Editor, overrides: Record<string, unknown> = {}) => {
    const onCommand = vi.fn()
    const view = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, { onCommand, comment: { id: null, done: false }, ...overrides }),
      ),
    )
    const byLabel = (label: string) =>
      [...view.container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find((b) =>
        b.querySelector('.ctx-label')?.textContent?.startsWith(label),
      )!
    return { ...view, onCommand, byLabel }
  }

  it('a caret in a cell gets the menu Word shows there, each item a shared command', () => {
    const { editor, cells } = tableEditor()
    editor.commands.setTextSelection(cells[0] + 2)
    const { container, unmount, onCommand, byLabel } = open(editor)
    act(() => {
      // React synthesizes onMouseEnter from bubbling mouseover
      byLabel('Insert').dispatchEvent(new MouseEvent('mouseover', { bubbles: true }))
      byLabel('Insert').dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }))
    })
    act(() => byLabel('Insert Above').click())
    expect(onCommand).toHaveBeenLastCalledWith('table-insert-rows-above')
    act(() => byLabel('Table Properties').click())
    expect(onCommand).toHaveBeenLastCalledWith('table-properties')
    act(() => byLabel('Delete Cells').click())
    expect(onCommand).toHaveBeenLastCalledWith('table-delete-cells')
    act(() => byLabel('Split Cells').click())
    expect(onCommand).toHaveBeenLastCalledWith('table-split-cells')
    expect(outline(container)).toEqual([
      'Cut',
      'Copy',
      'Paste',
      '---',
      'Reply To Comment',
      'Delete Comment',
      'Resolve Comment',
      '---',
      'Insert',
      'Delete Cells…',
      'Split Cells…',
      '---',
      'Borders and Shading…',
      'Text Direction…',
      'Cell Alignment',
      'Table Properties…',
      '---',
      'New Comment',
    ])
    unmount()
    editor.destroy()
  })

  it('the withholding item stays in a cell, in a group of its own before New Comment', () => {
    const { editor, cells } = tableEditor()
    editor.commands.setTextSelection(cells[0] + 2)
    const { container, unmount } = open(editor, { onRedact: noop })
    expect(outline(container).slice(-5)).toEqual([
      'Table Properties…',
      '---',
      'Hide the selection from AI',
      '---',
      'New Comment',
    ])
    unmount()
    editor.destroy()
  })

  it('a selected row swaps in Delete Rows, Merge Cells and Distribute, the comment items after the table', () => {
    const { editor, cells } = tableEditor()
    selectCells(editor, cells[0], cells[1])
    const { container, unmount, onCommand, byLabel } = open(editor)
    expect(outline(container)).toEqual([
      'Cut',
      'Copy',
      'Paste',
      '---',
      'Insert',
      'Delete Rows',
      'Merge Cells',
      'Distribute Rows Evenly',
      'Distribute Columns Evenly',
      '---',
      'Borders and Shading…',
      'Text Direction…',
      'Cell Alignment',
      'Table Properties…',
      '---',
      'Reply To Comment',
      'Delete Comment',
      'Resolve Comment',
      '---',
      'New Comment',
    ])
    act(() => byLabel('Delete Rows').click())
    expect(onCommand).toHaveBeenLastCalledWith('table-delete-rows')
    act(() => byLabel('Merge Cells').click())
    expect(onCommand).toHaveBeenLastCalledWith('table-merge-cells')
    act(() => byLabel('Distribute Rows Evenly').click())
    expect(onCommand).toHaveBeenLastCalledWith('table-distribute-rows')
    act(() => byLabel('Distribute Columns Evenly').click())
    expect(onCommand).toHaveBeenLastCalledWith('table-distribute-columns')
    unmount()
    editor.destroy()
  })

  it('a selected column offers Delete Columns, the whole table Delete Table, one cell alone Delete Cells\u2026', () => {
    const { editor, cells } = tableEditor()
    const deleteItem = () => {
      const { container, unmount } = open(editor)
      const rows = outline(container)
      unmount()
      return rows[rows.indexOf('Insert') + 1]
    }
    selectCells(editor, cells[0], cells[2])
    expect(deleteItem()).toBe('Delete Columns')
    selectCells(editor, cells[0], cells[3])
    expect(deleteItem()).toBe('Delete Table')
    selectCells(editor, cells[0], cells[0])
    expect(deleteItem()).toBe('Delete Cells…')
    editor.destroy()
  })

  it('a table selected as a whole, as its move handle selects it, gets Delete Table and a live Merge Cells', () => {
    const { editor } = tableEditor()
    let tablePos = -1
    editor.state.doc.descendants((node, pos) => {
      if (tablePos < 0 && node.type.name === 'docTable') tablePos = pos
    })
    editor.view.dispatch(
      editor.state.tr.setSelection(NodeSelection.create(editor.state.doc, tablePos)),
    )
    const { container, unmount, byLabel } = open(editor)
    const rows = outline(container)
    expect(rows.slice(rows.indexOf('Insert'), rows.indexOf('Borders and Shading…'))).toEqual([
      'Insert',
      'Delete Table',
      'Merge Cells',
      'Distribute Rows Evenly',
      'Distribute Columns Evenly',
      '---',
    ])
    expect(byLabel('Merge Cells').disabled).toBe(false)
    unmount()
    editor.destroy()
  })
})

describe('EditorContextMenu — submenu viewport fit', () => {
  const byLabel = (container: Element, label: string) =>
    [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
      (b) => b.querySelector('.ctx-label')?.textContent === label,
    )!

  function openSynonyms(container: HTMLElement) {
    Object.assign(window.desktop, {
      getAiSettings: async () => ({}),
      aiChat: () => new Promise(() => undefined),
    })
    const synonyms = byLabel(container, 'Synonyms')
    act(() => {
      // React synthesizes onMouseEnter from bubbling mouseover
      synonyms.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }))
      synonyms.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }))
    })
    return container.querySelector<HTMLDivElement>('.ctx-submenu')!
  }

  it('opens a submenu in the box that scrolls what does not fit', () => {
    const editor = createEditor()
    select(editor, 1, 4)
    const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
    const submenu = openSynonyms(container)
    expect(submenu).not.toBeNull()
    expect(submenu.classList.contains('ctx-submenu-scroll')).toBe(true)
    unmount()
    editor.destroy()
  })

  it('shifts a submenu up when the opening item sits near the window bottom', () => {
    const editor = createEditor()
    select(editor, 1, 4)
    const realRect = HTMLElement.prototype.getBoundingClientRect
    HTMLElement.prototype.getBoundingClientRect = function () {
      if (this instanceof Element && this.classList.contains('ctx-submenu')) {
        // opens 600px down, 200px tall: the bottom edge (800) passes the 768px window
        return {
          top: 600,
          bottom: 800,
          height: 200,
          left: 0,
          right: 200,
          width: 200,
          x: 0,
          y: 600,
          toJSON: () => ({}),
        } as DOMRect
      }
      return realRect.call(this)
    }
    try {
      const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
      const submenu = openSynonyms(container)
      // bottom-aligned to the window edge (760): top = 760 - 200 - 600 - 5
      expect(submenu.style.top).toBe('-45px')
      unmount()
    } finally {
      HTMLElement.prototype.getBoundingClientRect = realRect
    }
    editor.destroy()
  })

  it('fits the Synonyms submenu again when the candidates arrive after it opened', async () => {
    const editor = createEditor()
    select(editor, 1, 4)
    let reply: (value: { ok: boolean; content: string }) => void = () => undefined
    Object.assign(window.desktop, {
      getAiSettings: async () => ({}),
      aiChat: () => new Promise((resolve) => (reply = resolve)),
    })
    const realRect = HTMLElement.prototype.getBoundingClientRect
    HTMLElement.prototype.getBoundingClientRect = function () {
      if (this instanceof Element && this.classList.contains('ctx-submenu')) {
        // 700px down; two rows while loading (fits in 768), 40px per row once filled
        const rows = this.querySelectorAll('.ctx-item').length
        const height = rows <= 2 ? 50 : rows * 40
        const top = 700 + (parseFloat((this as HTMLElement).style.top) + 5 || 0)
        return {
          top,
          bottom: top + height,
          height,
          left: 0,
          right: 200,
          width: 200,
          x: 0,
          y: top,
          toJSON: () => ({}),
        } as DOMRect
      }
      return realRect.call(this)
    }
    try {
      const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
      const synonym = [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
        (b) => b.querySelector('.ctx-label')?.textContent === 'Synonyms',
      )!
      await act(async () => {
        synonym.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }))
      })
      const submenu = container.querySelector<HTMLElement>('.ctx-submenu')!
      expect(submenu.style.top).toBe('')
      await act(async () => reply({ ok: true, content: 'a\nb\nc\nd' }))
      // five rows, 200px: bottom-aligned to 760 from the wrap edge at 705
      expect(submenu.style.top).toBe('-145px')
      unmount()
    } finally {
      HTMLElement.prototype.getBoundingClientRect = realRect
    }
    editor.destroy()
  })

  it('anchors a taller-than-the-window submenu to the top and scrolls it', () => {
    const editor = createEditor()
    select(editor, 1, 4)
    const realRect = HTMLElement.prototype.getBoundingClientRect
    HTMLElement.prototype.getBoundingClientRect = function () {
      if (this instanceof Element && this.classList.contains('ctx-submenu')) {
        return {
          top: 500,
          bottom: 5000,
          height: 4500,
          left: 0,
          right: 200,
          width: 200,
          x: 0,
          y: 500,
          toJSON: () => ({}),
        } as DOMRect
      }
      return realRect.call(this)
    }
    try {
      const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
      const submenu = openSynonyms(container)
      // rising any further would poke out the top: anchored at the 8px margin,
      // the rest scrolls (maxHeight = window 768 - top margin 8 - bottom margin 8)
      expect(submenu.style.top).toBe('-497px')
      expect(submenu.style.maxHeight).toBe('752px')
      unmount()
    } finally {
      HTMLElement.prototype.getBoundingClientRect = realRect
    }
    editor.destroy()
  })
})

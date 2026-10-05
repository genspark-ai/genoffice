import { describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { createElement, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { act } from 'react'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { LocaleProvider, setModuleLang } from '../src/renderer/i18n/locale'
import { EditorContextMenu } from '../src/renderer/components/ContextMenu'
import { RedactDialog } from '../src/renderer/editor/RedactDialog'
import { modelTextOf, redactionCount } from '../src/renderer/ai/redact-view'

/**
 * The user-facing half of withholding a span: the right-click item that asks for
 * a label, and the dialog that shows the resulting marker before committing.
 *
 * The two halves are tested together at the end, because the point of the item
 * is what it does to the document, not that it renders.
 */

const LABEL = 'Hide the selection from AI'

function createEditor(): Editor {
  return new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: 'call the client about the order' }],
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
    onRedact: noop,
    ...overrides,
  }
}

const item = (container: Element, label: string) =>
  container.querySelector<HTMLButtonElement>(`.ctx-item:has(.ctx-label)`) &&
  [...container.querySelectorAll<HTMLButtonElement>('.ctx-item')].find(
    (b) => b.querySelector('.ctx-label')?.textContent === label,
  )

Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })
setModuleLang('en')

describe('RedactDialog', () => {
  function open(overrides: Partial<React.ComponentProps<typeof RedactDialog>> = {}) {
    const onSubmit = vi.fn()
    const onCancel = vi.fn()
    const view = render(
      createElement(RedactDialog, { seed: 'client phone', onSubmit, onCancel, ...overrides }),
    )
    const input = view.container.querySelector<HTMLInputElement>('input')!
    const primary = [...view.container.querySelectorAll('button')].find((b) =>
      b.classList.contains('primary'),
    )!
    return { ...view, input, primary, onSubmit, onCancel }
  }

  it('offers the selected text as a starting label', () => {
    const { input, unmount } = open()
    expect(input.value).toBe('client phone')
    unmount()
  })

  it('shows the marker the model will see, and follows what is typed', () => {
    // the reader must know what is about to be exposed before committing
    const { container, input, unmount } = open()
    const preview = () => container.querySelector('.redact-preview')!.textContent
    expect(preview()).toBe('{{client phone}}')
    act(() => {
      const setter = Object.getOwnPropertyDescriptor(
        window.HTMLInputElement.prototype,
        'value',
      )!.set!
      setter.call(input, 'order number')
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    expect(preview()).toBe('{{order number}}')
    unmount()
  })

  it('previews the neutral fallback for an empty label', () => {
    const { container, input, unmount } = open({ seed: '' })
    expect(container.querySelector('.redact-preview')!.textContent).toBe('{{private}}')
    act(() => {
      const setter = Object.getOwnPropertyDescriptor(
        window.HTMLInputElement.prototype,
        'value',
      )!.set!
      setter.call(input, '')
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    expect(container.querySelector('.redact-preview')!.textContent).toBe('{{private}}')
    unmount()
  })

  it('submits the sanitized label, dropping characters that would blur the marker', () => {
    const { input, primary, onSubmit, unmount } = open({ seed: '' })
    act(() => {
      const setter = Object.getOwnPropertyDescriptor(
        window.HTMLInputElement.prototype,
        'value',
      )!.set!
      setter.call(input, 'a{b}c')
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    act(() => primary.click())
    expect(onSubmit).toHaveBeenCalledWith('abc')
    unmount()
  })

  it('cannot be committed while the label is empty', () => {
    const { primary, onSubmit, unmount } = open({ seed: '' })
    expect(primary.disabled).toBe(true)
    act(() => primary.click())
    expect(onSubmit).not.toHaveBeenCalled()
    unmount()
  })

  it('submits on Enter and cancels on Escape', () => {
    const key = (input: HTMLInputElement, k: string) =>
      act(() => input.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true })))
    const enter = open()
    key(enter.input, 'Enter')
    expect(enter.onSubmit).toHaveBeenCalledWith('client phone')
    enter.unmount()

    const esc = open()
    key(esc.input, 'Escape')
    expect(esc.onCancel).toHaveBeenCalledOnce()
    expect(esc.onSubmit).not.toHaveBeenCalled()
    esc.unmount()
  })

  it('cancels without submitting when the backdrop is dismissed', () => {
    const { container, onSubmit, onCancel, unmount } = open()
    act(() => {
      const backdrop = container.querySelector<HTMLElement>('.modal-backdrop')!
      backdrop.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: 0, clientY: 0 }))
    })
    expect(onCancel).toHaveBeenCalledOnce()
    expect(onSubmit).not.toHaveBeenCalled()
    unmount()
  })
})

describe('EditorContextMenu — the withholding item', () => {
  it('appears once text is selected and routes to onRedact', () => {
    const editor = createEditor()
    select(editor, 1, 5)
    const onRedact = vi.fn()
    const onClose = vi.fn()
    const { container, unmount } = render(
      createElement(EditorContextMenu, menuProps(editor, { onRedact, onClose })),
    )
    const entry = item(container, LABEL)!
    expect(entry).toBeTruthy()
    expect(entry.disabled).toBe(false)
    act(() => entry.click())
    expect(onRedact).toHaveBeenCalledOnce()
    // the menu closes first, as every other item does
    expect(onClose).toHaveBeenCalled()
    unmount()
    editor.destroy()
  })

  it('sits with the AI entries, right after Translate', () => {
    const editor = createEditor()
    select(editor, 1, 5)
    const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
    const names = [...container.querySelectorAll('.ctx-label')].map((el) => el.textContent)
    const at = names.indexOf('Translate')
    expect(names[at + 1]).toBe(LABEL)
    unmount()
    editor.destroy()
  })

  it('is disabled, not hidden, with no selection, so the menu does not reflow', () => {
    const editor = createEditor()
    const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
    expect(item(container, LABEL)!.disabled).toBe(true)
    unmount()
    editor.destroy()
  })

  it('is disabled when the document cannot be edited', () => {
    const editor = createEditor()
    editor.setEditable(false)
    select(editor, 1, 5)
    const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
    expect(item(container, LABEL)!.disabled).toBe(true)
    unmount()
    editor.destroy()
  })

  it('carries no AI badge: the item withholds words, it does not spend them', () => {
    // the badge reads "Uses AI", which is the opposite of what this item does
    const editor = createEditor()
    select(editor, 1, 5)
    const { container, unmount } = render(createElement(EditorContextMenu, menuProps(editor)))
    expect(item(container, LABEL)!.querySelector('.copilot-badge')).toBeNull()
    unmount()
    editor.destroy()
  })

  it('is disabled for a picture, which cannot carry a mark through a save', () => {
    // A picture alone is a block node that takes no marks, and an inline one
    // loses the mark on save; offering the item would hide it for a session
    // and then drop the decision silently. Covered end to end in
    // redact-image-mark.test.ts.
    const editor = createEditor()
    const { container, unmount } = render(
      createElement(
        EditorContextMenu,
        menuProps(editor, { menu: { x: 10, y: 10, imageSrc: 'data:image/png;base64,AAAA' } }),
      ),
    )
    expect(item(container, LABEL)!.disabled).toBe(true)
    unmount()
    editor.destroy()
  })
})

describe('the redaction styles', () => {
  // jsdom never applies the stylesheet, so nothing above can catch a class
  // that was never written or a colour that ignores the theme.
  const css = readFileSync(join(__dirname, '..', 'src', 'renderer', 'styles.css'), 'utf8')
  const rule = (selector: string) => {
    const at = css.indexOf(`\n${selector} {`)
    expect(at, `${selector} is not defined in styles.css`).toBeGreaterThan(-1)
    return css.slice(at, css.indexOf('}', at))
  }

  it('defines every class the dialog and the mark render', () => {
    for (const selector of ['.redact-span', '.redact-desc', '.redact-preview']) {
      expect(rule(selector).length).toBeGreaterThan(0)
    }
  })

  it('draws the mark with theme variables, not colours of its own', () => {
    const span = rule('.redact-span')
    expect(span).toContain('var(--accent)')
    expect(span).not.toMatch(/#[0-9a-fA-F]{3,8}\b/)
    expect(span).not.toMatch(/\brgba?\(|\bhsla?\(/)
  })

  it('underlines the withheld span, so it reads differently from a highlight', () => {
    // every other decoration here is a fill or a style; a rule is what says
    // "the model cannot see this" at a glance
    expect(rule('.redact-span')).toMatch(/border-bottom:\s*2px solid/)
  })
})

describe('the item and the dialog together', () => {
  /** the App.tsx wiring, trimmed to the two pieces this feature needs */
  function Harness({ editor }: { editor: Editor }) {
    const [target, setTarget] = useState<{ from: number; to: number; seed: string } | null>(null)
    return createElement(
      'div',
      null,
      createElement(EditorContextMenu, {
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
        onRedact: () => {
          const { from, to } = editor.state.selection
          if (from === to) return
          setTarget({ from, to, seed: editor.state.doc.textBetween(from, to, ' ').trim() })
        },
      }),
      target &&
        createElement(RedactDialog, {
          seed: target.seed,
          onCancel: () => setTarget(null),
          onSubmit: (label: string) => {
            editor
              .chain()
              .setTextSelection({ from: target.from, to: target.to })
              .setRedaction(label)
              .run()
            setTarget(null)
          },
        }),
    )
  }

  it('marks the selection and hands the model a marker, keeping the words', () => {
    const editor = createEditor()
    select(editor, 10, 16) // "client"
    const { container, unmount } = render(createElement(Harness, { editor }))

    act(() => item(container, LABEL)!.click())
    const input = container.querySelector<HTMLInputElement>('input')!
    expect(input.value).toBe('client')
    act(() => {
      const setter = Object.getOwnPropertyDescriptor(
        window.HTMLInputElement.prototype,
        'value',
      )!.set!
      setter.call(input, 'client')
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
    act(() =>
      [...container.querySelectorAll('button')]
        .find((b) => b.classList.contains('primary'))!
        .click(),
    )

    // the document still holds the reader's own words
    expect(editor.state.doc.textContent).toBe('call the client about the order')
    expect(redactionCount(editor.getJSON() as never)).toBe(1)
    // and the model's view is what changed
    expect(modelTextOf(editor.getJSON() as never)).toBe('call the {{client}} about the order')
    // the mark is visible in the editor
    expect(editor.view.dom.querySelector('.redact-span')).not.toBeNull()
    // the dialog is gone
    expect(container.querySelector('.modal-backdrop')).toBeNull()

    unmount()
    editor.destroy()
  })

  it('leaves the document untouched when the reader cancels', () => {
    const editor = createEditor()
    select(editor, 10, 16)
    const { container, unmount } = render(createElement(Harness, { editor }))
    act(() => item(container, LABEL)!.click())
    act(() =>
      [...container.querySelectorAll('button')].find((b) => b.textContent === 'Cancel')!.click(),
    )
    expect(redactionCount(editor.getJSON() as never)).toBe(0)
    expect(editor.state.doc.textContent).toBe('call the client about the order')
    expect(editor.getHTML()).not.toContain('redact-span')
    unmount()
    editor.destroy()
  })
})

// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { ICommandService } from '@univerjs/core'

import { RedactDialog } from '../src/renderer/components/RedactDialog'
import { HIDE_SELECTION_COMMAND_ID, installRedactMenu } from '../src/renderer/redact-menu'
import { strings } from '../src/renderer/i18n/strings'
import { hideFromAiLocale } from '../src/renderer/univer-locales'

/**
 * The reader-facing surface: the dialog that asks what withheld cells stand
 * for, and the grid menu item that opens it.
 *
 * Both have a job a render check cannot fake. The dialog must hand back the
 * label the reader *typed* rather than the seed it was offered — confirming
 * the seed instead of the edit looks identical on screen. The menu command
 * must report the exact rectangle that was selected, because a mark on the
 * wrong cells is worse than no mark: it hides nothing and leaves the reader
 * believing something is hidden that is not.
 *
 * Written with `createElement` rather than JSX because this package's vitest
 * include pattern only picks up `.test.ts`, so a `.tsx` here would not run.
 */

beforeEach(() => vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true))
const cleanups: Array<() => void> = []
afterEach(() => {
  cleanups.splice(0).forEach((cleanup) => cleanup())
  vi.unstubAllGlobals()
})

type DialogProps = {
  seed: string
  rangeLabel: string
  onSubmit: (label: string) => void
  onCancel: () => void
}

function dialog(props: Partial<DialogProps> = {}) {
  return createElement(RedactDialog, {
    seed: '',
    rangeLabel: 'Customers!B2',
    onSubmit: () => {},
    onCancel: () => {},
    ...props,
  } as DialogProps)
}

function mount(node: ReturnType<typeof createElement>) {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root: Root = createRoot(container)
  act(() => root.render(node))
  cleanups.push(() => {
    act(() => root.unmount())
    container.remove()
  })
  return container
}

function type(input: HTMLInputElement, text: string) {
  act(() => {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set
    setter?.call(input, text)
    input.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

function click(element: Element) {
  act(() => element.dispatchEvent(new MouseEvent('click', { bubbles: true })))
}

describe('the redaction dialog', () => {
  it('offers the cell text as a starting label, selected so typing replaces it', () => {
    const input = mount(dialog({ seed: '13800138000' })).querySelector('input')!
    expect(input.value).toBe('13800138000')
    // Pre-selected: the reader types over the phone number rather than
    // appending to it. A seed that is not selected is a very easy mistake to
    // make here, and the label is the only thing the model will ever see.
    expect(input.selectionStart).toBe(0)
    expect(input.selectionEnd).toBe('13800138000'.length)
  })

  it('hands back what the reader typed, not the seed', () => {
    const onSubmit = vi.fn()
    const container = mount(dialog({ seed: '13800138000', onSubmit }))
    type(container.querySelector('input')!, 'client phone')
    click(container.querySelector('.btn-primary')!)
    expect(onSubmit).toHaveBeenCalledWith('client phone')
  })

  it('submits the cleaned label, not the raw keystrokes', () => {
    // The seed and the edit look identical once the reader types something
    // already clean, so a dialog that confirmed the raw field would pass every
    // other test here. A padded label is what tells the two apart — and the
    // whitespace matters, because the label is the literal the model reads.
    const onSubmit = vi.fn()
    const container = mount(dialog({ onSubmit }))
    type(container.querySelector('input')!, '  client phone  ')
    click(container.querySelector('.btn-primary')!)
    expect(onSubmit).toHaveBeenCalledWith('client phone')
  })

  it('strips the characters that would break a placeholder', () => {
    const onSubmit = vi.fn()
    const container = mount(dialog({ onSubmit }))
    type(container.querySelector('input')!, 'a{b}c"d')
    click(container.querySelector('.btn-primary')!)
    expect(onSubmit).toHaveBeenCalledWith('abcd')
  })

  it('shows the placeholder the model will read, live', () => {
    const container = mount(dialog({ rangeLabel: 'Customers!B2:B500' }))
    expect(container.querySelector('.redact-dialog-preview')!.textContent).toBe('{{private}}')
    type(container.querySelector('input')!, 'client phone')
    expect(container.querySelector('.redact-dialog-preview')!.textContent).toBe('{{client phone}}')
  })

  it('names the range so a whole-column mark is never a surprise', () => {
    const container = mount(dialog({ rangeLabel: 'Customers!B2:B500' }))
    expect(container.querySelector('.redact-dialog-range')!.textContent).toBe('Customers!B2:B500')
  })

  it('cannot be confirmed without a usable label', () => {
    const onSubmit = vi.fn()
    const container = mount(dialog({ onSubmit }))
    const confirm = () => container.querySelector<HTMLButtonElement>('.btn-primary')!
    expect(confirm().disabled).toBe(true)
    type(container.querySelector('input')!, '   ')
    expect(confirm().disabled).toBe(true)
    expect(onSubmit).not.toHaveBeenCalled()
  })

  it('cancels on Escape and on a click outside the card', () => {
    const onCancel = vi.fn()
    const container = mount(dialog({ onCancel }))
    act(() => {
      container
        .querySelector('input')!
        .dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
    })
    expect(onCancel).toHaveBeenCalledTimes(1)
    act(() => {
      container
        .querySelector('.redact-dialog')!
        .dispatchEvent(new MouseEvent('mousedown', { bubbles: true }))
    })
    expect(onCancel).toHaveBeenCalledTimes(2)
  })
})

describe('the grid menu item', () => {
  /// A runtime whose injector answers only the two services the installer asks
  /// for, matched on the real tokens so the shipped wiring is what runs.
  function harness(range: unknown) {
    const merged: unknown[] = []
    let fire: (() => boolean) | null = null
    const injector = {
      get: (token: unknown) => {
        if (token === ICommandService) {
          return {
            registerCommand: (command: { handler: () => boolean }) => {
              fire = () => command.handler()
              return { dispose: () => {} }
            },
          }
        }
        return { mergeMenu: (schema: unknown) => merged.push(schema) }
      },
    }
    return {
      merged,
      fire: () => fire?.() ?? false,
      runtime: {
        univer: { __getInjector: () => injector },
        univerAPI: {
          getActiveWorkbook: () => ({
            getActiveSheet: () => ({
              getSheetId: () => 'sh1',
              getActiveRange: () => ({ getRange: () => range }),
            }),
          }),
        },
      } as never,
    }
  }

  it('reports the exact rectangle that was selected', () => {
    const onRequest = vi.fn()
    const h = harness({ startRow: 1, endRow: 9, startColumn: 1, endColumn: 1 })
    installRedactMenu(h.runtime, onRequest).dispose()
    expect(h.fire()).toBe(true)
    expect(onRequest).toHaveBeenCalledWith({
      sheetId: 'sh1',
      startRow: 1,
      endRow: 9,
      startColumn: 1,
      endColumn: 1,
      isSingleCell: false,
    })
  })

  it('tells a single cell apart from a range', () => {
    const onRequest = vi.fn()
    const h = harness({ startRow: 3, endRow: 3, startColumn: 2, endColumn: 2 })
    installRedactMenu(h.runtime, onRequest).dispose()
    h.fire()
    expect(onRequest.mock.calls[0]?.[0]?.isSingleCell).toBe(true)
  })

  it('does nothing when the right-click left no cell range', () => {
    // The click landed on a header or the formula bar. Guessing a target here
    // would mark cells the reader never looked at.
    const onRequest = vi.fn()
    const h = harness(null)
    installRedactMenu(h.runtime, onRequest).dispose()
    expect(h.fire()).toBe(false)
    expect(onRequest).not.toHaveBeenCalled()
  })

  it('contributes the item to the grid context menu', () => {
    const h = harness({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 })
    installRedactMenu(h.runtime, vi.fn()).dispose()
    expect(h.merged).toHaveLength(1)
    // Keyed by the command it runs, which is what Univer invokes on press.
    expect(JSON.stringify(h.merged)).toContain(HIDE_SELECTION_COMMAND_ID)
  })
})

describe('the menu label reaches Univer as a locale key', () => {
  // Univer resolves a menu item's `title` through its own LocaleService, so a
  // missing key does not fail loudly — the item simply renders untranslated.
  function titleFor(lang: 'en' | 'ja' | 'th') {
    const pack = hideFromAiLocale({ 'sheets-ui': { rightClick: { cut: 'Cut' } } }, lang)
    return (pack['sheets-ui'] as { rightClick: Record<string, string> }).rightClick
      .hideSelectionFromAi
  }

  it('reads the wording out of the app string, for every language it is given', () => {
    // Compared against the app's own table rather than a literal, so the
    // assertion holds for all 21 languages and cannot rot when one is edited.
    // en is what the runtime boots with and ja is a Univer-pack language;
    // both must come from here, which is what keeps the gesture identical
    // across the other editors.
    expect(titleFor('en')).toBe(strings.en.redactMenuLabel)
    expect(titleFor('ja')).toBe(strings.ja.redactMenuLabel)
    // ...and ja is genuinely translated, not quietly falling back to English.
    expect(strings.ja.redactMenuLabel).not.toBe(strings.en.redactMenuLabel)
  })

  it('leaves the pack it merged into untouched', () => {
    const pack = hideFromAiLocale({ 'sheets-ui': { rightClick: { cut: 'Cut' } } }, 'en')
    const rightClick = (pack['sheets-ui'] as { rightClick: Record<string, string> }).rightClick
    expect(rightClick.cut).toBe('Cut')
  })

  it('stays English where Univer has no pack, like every other right-click item', () => {
    // th/nl/ms/he/hi/cs/zh-TW keep the English pack the runtime booted with,
    // where every other entry is English. The redaction item follows that
    // rather than being the one string in the menu with a language of its own
    // — so it must not fall back to the app's Thai wording.
    expect(titleFor('th')).toBe(strings.en.redactMenuLabel)
    expect(titleFor('th')).not.toBe(strings.th.redactMenuLabel)
  })
})

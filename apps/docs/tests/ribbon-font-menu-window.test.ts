import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { Ribbon } from '../src/renderer/components/Ribbon'
import { ribbonProps } from './helpers/ribbon-props'

const SYSTEM_FAMILIES = Array.from(
  { length: 600 },
  (_, i) => `Sys Face ${String(i).padStart(3, '0')}`,
)

describe('font family menu with hundreds of system fonts (genoffice#2211)', () => {
  let editor: Editor
  let host: HTMLDivElement
  let root: Root
  let resolveFonts: (fonts: { family: string }[]) => void = () => {}
  const render = () =>
    act(() => root.render(createElement(Ribbon, ribbonProps(editor, computeFormatState(editor)))))
  const menu = () => host.querySelector<HTMLElement>('.rb-font-family-menu')
  const rows = () =>
    Array.from(menu()?.querySelectorAll<HTMLButtonElement>('button[data-font-key]') ?? [])
  const previewed = () => rows().filter((b) => b.style.fontFamily !== '')

  beforeEach(() => {
    vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
    vi.stubGlobal(
      'ResizeObserver',
      class {
        observe() {}
        disconnect() {}
        unobserve() {}
      },
    )
    vi.stubGlobal(
      'queryLocalFonts',
      () =>
        new Promise<{ family: string }[]>((resolve) => {
          resolveFonts = resolve
        }),
    )
    editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'x' }] }],
      },
    })
    editor.commands.selectAll()
    host = document.createElement('div')
    document.body.append(host)
    root = createRoot(host)
    render()
  })
  afterEach(() => {
    act(() => root.unmount())
    host.remove()
    editor.destroy()
    vi.unstubAllGlobals()
  })

  it('opens with the built-in section at once and previews only rows near the viewport', async () => {
    const caret = host.querySelector<HTMLButtonElement>('.rb-font-group .rb-combo-caret')!
    act(() => caret.click())
    expect(menu()).not.toBeNull()
    const builtin = rows().length
    expect(builtin).toBeGreaterThan(0)
    expect(rows().every((b) => b.dataset.fontKey!.startsWith('b:'))).toBe(true)

    await act(async () => {
      resolveFonts(SYSTEM_FAMILIES.map((family) => ({ family })))
      await Promise.resolve()
    })
    expect(rows().length).toBe(builtin + SYSTEM_FAMILIES.length)
    expect(previewed().length).toBeGreaterThan(0)
    expect(previewed().length).toBeLessThanOrEqual(40)
    expect(previewed()[0]).toBe(rows()[0])

    // scrolling moves the preview window instead of widening it
    const list = menu()!
    Object.defineProperty(list, 'scrollTop', { value: 26 * 400, configurable: true })
    await act(async () => {
      list.dispatchEvent(new Event('scroll'))
      await new Promise((r) => requestAnimationFrame(r))
    })
    const after = previewed()
    expect(after.length).toBeLessThanOrEqual(40)
    expect(after[0]).not.toBe(rows()[0])
    expect(after[0]!.dataset.fontKey!.startsWith('s:')).toBe(true)
  })
})

import { describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { LocaleProvider, setModuleLang } from '../src/renderer/i18n/locale'
import { Ribbon, docModeOf } from '../src/renderer/components/Ribbon'
import { computeFormatState } from '../src/renderer/components/ribbon-format-state'
import { ribbonProps } from './helpers/ribbon-props'

Object.assign(window, { desktop: { onLanguageChanged: () => () => undefined } })
setModuleLang('en')

describe('docModeOf', () => {
  it('maps read mode and track changes onto the Word mode names', () => {
    expect(docModeOf(false, false)).toBe('editing')
    expect(docModeOf(false, true)).toBe('reviewing')
    // Viewing wins: read mode hides the recorder state
    expect(docModeOf(true, true)).toBe('viewing')
    expect(docModeOf(true, false)).toBe('viewing')
  })
})

describe('top-right mode cluster', () => {
  function mount(extra: Record<string, unknown>) {
    vi.stubGlobal(
      'ResizeObserver',
      class {
        observe() {}
        disconnect() {}
        unobserve() {}
      },
    )
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: {
        type: 'doc',
        content: [{ type: 'docParagraph', content: [{ type: 'text', text: 'hello' }] }],
      },
    })
    const host = document.createElement('div')
    document.body.append(host)
    const root = createRoot(host)
    const props = { ...ribbonProps(editor, computeFormatState(editor)), ...extra }
    act(() =>
      root.render(
        createElement(LocaleProvider, {
          initial: 'en',
          children: createElement(Ribbon, props as never),
        }),
      ),
    )
    return {
      host,
      cleanup: () => {
        act(() => root.unmount())
        host.remove()
        editor.destroy()
      },
    }
  }

  it('toggles the comments panel and switches modes through the dropdown', () => {
    const onToggleComments = vi.fn()
    const onTrackChanges = vi.fn()
    const onReadMode = vi.fn()
    const { host, cleanup } = mount({ onToggleComments, onTrackChanges, onReadMode })

    const cluster = host.querySelector('.ribbon-mode-cluster')!
    expect(cluster).toBeTruthy()
    const btns = [...cluster.querySelectorAll<HTMLButtonElement>('.ribbon-mode-btn')]
    const commentsBtn = btns[0]
    const modeBtn = btns[btns.length - 1]
    act(() => commentsBtn.click())
    expect(onToggleComments).toHaveBeenCalledOnce()

    expect(modeBtn.textContent).toContain('Editing')
    act(() => modeBtn.click())
    const items = [...host.querySelectorAll<HTMLButtonElement>('.mode-menu button')]
    expect(items.map((b) => b.textContent)).toEqual(['✓Editing', 'Reviewing', 'Viewing'])
    act(() => items[1].click())
    expect(onTrackChanges).toHaveBeenCalledWith(true)
    act(() => modeBtn.click())
    act(() => [...host.querySelectorAll<HTMLButtonElement>('.mode-menu button')][2].click())
    expect(onReadMode).toHaveBeenCalledWith(true)
    cleanup()
  })

  it('shows Viewing and leaves read mode when Editing is picked', () => {
    const onReadMode = vi.fn()
    const onTrackChanges = vi.fn()
    const { host, cleanup } = mount({ readMode: true, onReadMode, onTrackChanges })
    const modeBtn = [...host.querySelectorAll<HTMLButtonElement>('.ribbon-mode-btn')].at(-1)!
    expect(modeBtn.textContent).toContain('Viewing')
    act(() => modeBtn.click())
    act(() => [...host.querySelectorAll<HTMLButtonElement>('.mode-menu button')][0].click())
    expect(onReadMode).toHaveBeenCalledWith(false)
    expect(onTrackChanges).not.toHaveBeenCalled()
    cleanup()
  })

  it('disables Editing under forced track changes', () => {
    const { host, cleanup } = mount({ trackChanges: true, trackChangesForced: true })
    const modeBtn = [...host.querySelectorAll<HTMLButtonElement>('.ribbon-mode-btn')].at(-1)!
    expect(modeBtn.textContent).toContain('Reviewing')
    act(() => modeBtn.click())
    const items = [...host.querySelectorAll<HTMLButtonElement>('.mode-menu button')]
    expect(items[0].disabled).toBe(true)
    cleanup()
  })
})

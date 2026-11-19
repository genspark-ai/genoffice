import { act, createElement, type ReactElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'

vi.mock('@tiptap/react', async () => {
  const react = await import('react')
  return {
    NodeViewWrapper: ({
      children,
      className,
    }: {
      children: import('react').ReactNode
      className?: string
    }) => react.createElement('div', { className }, children),
    NodeViewContent: () => react.createElement('code'),
  }
})

import { AiPanel, type MarkdownAiDeps } from '../src/renderer/ai/AiPanel'
import { CodeBlockView } from '../src/renderer/editor/CodeBlockView'

const mountedRoots: Array<{ root: Root; container: HTMLElement }> = []

function mount(element: ReactElement): { root: Root; container: HTMLElement } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  mountedRoots.push({ root, container })
  return { root, container }
}

function unmount(root: Root): void {
  const mounted = mountedRoots.find((entry) => entry.root === root)
  act(() => root.unmount())
  mounted?.container.remove()
  if (mounted) mountedRoots.splice(mountedRoots.indexOf(mounted), 1)
}

function typeInto(textarea: HTMLTextAreaElement, text: string): void {
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
  act(() => {
    setter.call(textarea, text)
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
  })
}

beforeAll(() => {
  Element.prototype.scrollTo ??= () => {}
})

afterEach(() => {
  for (const { root, container } of mountedRoots.splice(0)) {
    act(() => root.unmount())
    container.remove()
  }
  vi.useRealTimers()
  vi.restoreAllMocks()
})

describe('CodeBlockView teardown', () => {
  it('clears copy feedback timeout when its node view unmounts', async () => {
    vi.useFakeTimers()
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { writeText: vi.fn(async () => {}) },
    })
    const clearTimeout = vi.spyOn(window, 'clearTimeout')
    const { root, container } = mount(
      createElement(CodeBlockView, {
        node: { attrs: { language: null }, textContent: 'const answer = 42' },
        updateAttributes: vi.fn(),
        editor: { isEditable: true },
      } as never),
    )

    await act(async () => {
      container.querySelector<HTMLButtonElement>('.md-codeblock-copy')!.click()
      await Promise.resolve()
    })
    unmount(root)

    expect(clearTimeout).toHaveBeenCalled()
  })
})

describe('AiPanel teardown', () => {
  it('cancels an in-flight IPC stream when the panel/tab unmounts', async () => {
    vi.useFakeTimers()
    const aiStream = vi.fn(async () => {})
    const aiStreamCancel = vi.fn(async () => {})
    const onAiStream = vi.fn(() => () => {})
    Object.defineProperty(window, 'markdownApi', {
      configurable: true,
      value: {
        getAiSettings: vi.fn(async () => ({})),
        aiStream,
        aiStreamCancel,
        onAiStream,
      },
    })
    Object.defineProperty(window, 'projectApi', { configurable: true, value: undefined })
    const deps: MarkdownAiDeps = {
      getEditor: () => null,
      getSnapshot: () => '',
      restoreSnapshot: () => {},
      onRunDone: () => {},
    }
    const { root, container } = mount(
      createElement(AiPanel, { deps, filePath: null, onCollapse: () => {} }),
    )
    const textarea = container.querySelector<HTMLTextAreaElement>('textarea')!
    typeInto(textarea, 'keep streaming')

    await act(async () => {
      container.querySelector<HTMLButtonElement>('.ai-send-btn')!.click()
      await Promise.resolve()
      await Promise.resolve()
    })
    expect(aiStream).toHaveBeenCalledTimes(1)

    unmount(root)

    expect(aiStreamCancel).toHaveBeenCalledTimes(1)
  })
})

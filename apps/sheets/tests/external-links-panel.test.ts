// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'

import { ExternalLinksPanel } from '../src/renderer/ExternalLinksPanel'

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
  document.body.innerHTML = ''
})

async function mount(element: React.ReactElement): Promise<HTMLDivElement> {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(element)
    await Promise.resolve()
  })
  return container
}

function click(element: Element): void {
  act(() => element.dispatchEvent(new MouseEvent('click', { bubbles: true })))
}

const LINKS = [
  {
    index: 1,
    rawTarget: 'gone.xlsx',
    source: null,
    reachable: false,
    sheetNames: ['Sheet1'],
  },
  {
    index: 2,
    rawTarget: 'source.xlsx',
    source: '/tmp/data/source.xlsx',
    reachable: true,
    sheetNames: ['Sheet1', 'Extra'],
  },
]

describe('external links panel', () => {
  it('lists every link with its status and the resolved source', async () => {
    const host = await mount(
      createElement(ExternalLinksPanel, {
        links: LINKS,
        busy: false,
        onUpdate: () => {},
        onReveal: () => {},
        onClose: () => {},
      }),
    )
    expect(host.textContent).toContain('gone.xlsx')
    expect(host.textContent).toContain('/tmp/data/source.xlsx')
    expect(host.textContent).toContain('已断开')
    expect(host.textContent).toContain('可用')
    // locate only on the reachable link
    expect(host.querySelectorAll('[data-tip="定位源文件"]')).toHaveLength(1)
  })

  it('an empty workbook shows the no-links note', async () => {
    const host = await mount(
      createElement(ExternalLinksPanel, {
        links: [],
        busy: false,
        onUpdate: () => {},
        onReveal: () => {},
        onClose: () => {},
      }),
    )
    expect(host.textContent).toContain('没有指向其他工作簿的链接')
  })

  it('a busy panel ignores update clicks', async () => {
    const onUpdate = vi.fn()
    const host = await mount(
      createElement(ExternalLinksPanel, {
        links: LINKS,
        busy: true,
        onUpdate,
        onReveal: () => {},
        onClose: () => {},
      }),
    )
    const update = [...host.querySelectorAll('button')].find(
      (button) => button.textContent === '更新值',
    )!
    expect(update.hasAttribute('disabled')).toBe(true)
    click(update)
    expect(onUpdate).not.toHaveBeenCalled()
  })

  it('a click on update reaches the handler', async () => {
    const onUpdate = vi.fn()
    const host = await mount(
      createElement(ExternalLinksPanel, {
        links: LINKS,
        busy: false,
        onUpdate,
        onReveal: () => {},
        onClose: () => {},
      }),
    )
    const update = [...host.querySelectorAll('button')].find(
      (button) => button.textContent === '更新值',
    )!
    click(update)
    expect(onUpdate).toHaveBeenCalledTimes(1)
  })
})

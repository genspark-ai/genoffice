// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'

import { ExternalLinksPanel } from '../src/renderer/ExternalLinksPanel'
import type { ExternalLinkRow } from '../src/shared/desktop-api'

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

/// What read-on-open produces: identity fields only, no reachability verdict.
const LISTED: ExternalLinkRow[] = [
  {
    index: 1,
    rawTarget: 'gone.xlsx',
    source: null,
    sheetNames: ['Sheet1'],
  },
  {
    index: 2,
    rawTarget: 'source.xlsx',
    source: '/tmp/data/source.xlsx',
    sheetNames: ['Sheet1', 'Extra'],
  },
]

/// What the update report produces: the same rows with reachability answered.
const UPDATED: ExternalLinkRow[] = [
  { ...LISTED[0]!, reachable: false },
  { ...LISTED[1]!, reachable: true },
]

function panel(links: readonly ExternalLinkRow[], overrides: Record<string, unknown> = {}) {
  return createElement(ExternalLinksPanel, {
    links,
    busy: false,
    onUpdate: () => {},
    onReveal: () => {},
    onClose: () => {},
    ...overrides,
  })
}

describe('external links panel', () => {
  it('lists the links on open with reachability unknown', async () => {
    const host = await mount(panel(LISTED))
    expect(host.textContent).toContain('gone.xlsx')
    expect(host.textContent).toContain('/tmp/data/source.xlsx')
    expect(host.textContent).toContain('Sheet1, Extra')
    // read-on-open never probes the targets: the status is unknown until Update
    expect(host.textContent).toContain('—')
    expect(host.textContent).not.toContain('可用')
    expect(host.textContent).not.toContain('已断开')
    expect(host.querySelectorAll('[data-tip="定位源文件"]')).toHaveLength(0)
  })

  it('shows the reachability the update answered', async () => {
    const host = await mount(panel(UPDATED))
    expect(host.textContent).toContain('可用')
    expect(host.textContent).toContain('已断开')
    // locate only on the reachable link
    expect(host.querySelectorAll('[data-tip="定位源文件"]')).toHaveLength(1)
  })

  it('an empty workbook shows the no-links note', async () => {
    const host = await mount(panel([]))
    expect(host.textContent).toContain('没有指向其他工作簿的链接')
  })

  it('a busy panel ignores update clicks', async () => {
    const onUpdate = vi.fn()
    const host = await mount(panel(LISTED, { busy: true, onUpdate }))
    const update = [...host.querySelectorAll('button')].find(
      (button) => button.textContent === '更新值',
    )!
    expect(update.hasAttribute('disabled')).toBe(true)
    click(update)
    expect(onUpdate).not.toHaveBeenCalled()
  })

  it('a click on update reaches the handler', async () => {
    const onUpdate = vi.fn()
    const host = await mount(panel(LISTED, { onUpdate }))
    const update = [...host.querySelectorAll('button')].find(
      (button) => button.textContent === '更新值',
    )!
    click(update)
    expect(onUpdate).toHaveBeenCalledTimes(1)
  })

  it('a reachable row locates its source', async () => {
    const onReveal = vi.fn()
    const host = await mount(panel(UPDATED, { onReveal }))
    const locate = host.querySelector('[data-tip="定位源文件"]')!
    click(locate)
    expect(onReveal).toHaveBeenCalledWith('/tmp/data/source.xlsx')
  })
})

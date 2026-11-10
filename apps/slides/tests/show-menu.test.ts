import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { ContextMenu } from '../src/renderer/components/ContextMenu'
import { t } from '../src/renderer/i18n/locale'
import { buildShowMenu, type ShowMenuState } from '../src/renderer/show-menu'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const actions = {
  next: vi.fn(),
  prev: vi.fn(),
  lastViewed: vi.fn(),
  seeAll: vi.fn(),
  setScreen: vi.fn(),
  end: vi.fn(),
}

const base: ShowMenuState = {
  pos: 1,
  count: 3,
  ended: false,
  pending: false,
  hasLastViewed: true,
  screen: 'none',
}

const find = (s: Partial<ShowMenuState>, key: Parameters<typeof t>[0]) =>
  buildShowMenu(t, { ...base, ...s }, actions).find((i) => i?.label === t(key))

describe('slide show context menu items', () => {
  it('lists PowerPoint items in order', () => {
    const labels = buildShowMenu(t, base, actions).map((i) => (i ? i.label : '-'))
    expect(labels).toEqual([
      t('paneShowMenuNext'),
      t('paneShowMenuPrev'),
      t('paneShowMenuLastViewed'),
      t('paneShowMenuSeeAll'),
      '-',
      t('paneShowMenuScreen'),
      '-',
      t('paneShowMenuEnd'),
    ])
  })

  it('disables Previous on the first slide and Next on the last', () => {
    expect(find({ pos: 0 }, 'paneShowMenuPrev')?.disabled).toBe(true)
    expect(find({ pos: 1 }, 'paneShowMenuPrev')?.disabled).toBe(false)
    expect(find({ pos: 2 }, 'paneShowMenuNext')?.disabled).toBe(true)
    expect(find({ pos: 1 }, 'paneShowMenuNext')?.disabled).toBe(false)
  })

  it('keeps Next enabled on the last slide while animation steps remain', () => {
    expect(find({ pos: 2, pending: true }, 'paneShowMenuNext')?.disabled).toBe(false)
  })

  it('on the end screen only Previous leads back', () => {
    expect(find({ pos: 2, ended: true }, 'paneShowMenuNext')?.disabled).toBe(true)
    expect(find({ pos: 2, ended: true }, 'paneShowMenuPrev')?.disabled).toBe(false)
  })

  it('disables Last Viewed until a slide was left behind', () => {
    expect(find({ hasLastViewed: false }, 'paneShowMenuLastViewed')?.disabled).toBe(true)
  })

  it('checks the active screen in the Screen flyout', () => {
    const sub = find({ screen: 'white' }, 'paneShowMenuScreen')?.sub ?? []
    expect(sub.map((i) => i?.checked)).toEqual([false, true])
    sub[1]?.onClick?.()
    expect(actions.setScreen).toHaveBeenCalledWith('white')
  })
})

describe('slide show context menu rendering', () => {
  const roots: Array<{ root: Root; container: HTMLElement }> = []
  afterEach(() => {
    for (const { root, container } of roots.splice(0)) {
      act(() => root.unmount())
      container.remove()
    }
  })

  function mount(s: Partial<ShowMenuState>) {
    const container = document.createElement('div')
    document.body.appendChild(container)
    const root = createRoot(container)
    const items = buildShowMenu(t, { ...base, ...s }, actions)
    act(() => root.render(createElement(ContextMenu, { x: 0, y: 0, items, onClose: () => {} })))
    roots.push({ root, container })
    return container
  }

  const button = (c: HTMLElement, label: string) =>
    Array.from(c.querySelectorAll<HTMLButtonElement>('.ctx-item')).find(
      (b) => b.textContent === label,
    )

  it('renders disabled buttons for the first slide', () => {
    const c = mount({ pos: 0, hasLastViewed: false })
    expect(button(c, t('paneShowMenuPrev'))?.disabled).toBe(true)
    expect(button(c, t('paneShowMenuLastViewed'))?.disabled).toBe(true)
    expect(button(c, t('paneShowMenuNext'))?.disabled).toBe(false)
    expect(button(c, t('paneShowMenuEnd'))?.disabled).toBe(false)
  })

  it('opens the Screen flyout on hover', () => {
    const c = mount({})
    expect(c.querySelector('.ctx-submenu')).toBeNull()
    const host = c.querySelector('.ctx-sub-host') as HTMLElement
    act(() => {
      host.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }))
    })
    const sub = c.querySelector('.ctx-submenu') as HTMLElement
    expect(sub).not.toBeNull()
    expect(sub.querySelectorAll('.ctx-item')).toHaveLength(2)
  })
})

// Word for Mac model: the selected tab collapses, any tab expands (and stays
// expanded), the collapsed tab row has no selected tab.
import { beforeEach, describe, expect, it } from 'vitest'
import { act, createElement, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { isRibbonToggleShortcut, readRibbonCollapsed, useRibbonCollapse } from '@genoffice/ui'

const TABS = ['home', 'insert'] as const
const LABELS = { collapse: 'Collapse', expand: 'Expand' }

function Ribbon() {
  const collapse = useRibbonCollapse('t.ribbon', LABELS)
  const [tab, setTab] = useState<string>('home')
  return createElement(
    'div',
    { className: collapse.rootClass },
    createElement(
      'div',
      { className: 'ribbon-tabs', onDoubleClick: collapse.onTabsDoubleClick },
      ...TABS.map((name) =>
        createElement(
          'button',
          {
            key: name,
            id: name,
            className: `ribbon-tab ${collapse.tabClass(tab === name)}`,
            'data-tip': collapse.tabTip(tab === name),
            onClick: () => {
              collapse.onTabPress(tab === name)
              setTab(name)
            },
          },
          name,
        ),
      ),
    ),
    createElement('div', { 'data-ribbon-body': '' }),
  )
}

const $ = (id: string) => document.getElementById(id) as HTMLButtonElement
const rootEl = () => document.querySelector('.ribbon-collapsible') as HTMLElement
const collapsed = () => rootEl().classList.contains('ribbon-collapsed')
const click = (id: string) => act(() => $(id).click())

describe('ribbon collapse', () => {
  beforeEach(() => {
    localStorage.removeItem('t.ribbon')
    document.body.innerHTML = ''
    const host = document.createElement('div')
    document.body.append(host)
    act(() => createRoot(host).render(createElement(Ribbon)))
  })

  it('pressing the selected tab collapses; pressing another tab only switches', () => {
    click('insert')
    expect(collapsed()).toBe(false)
    expect($('insert').className).toContain('active')
    click('insert')
    expect(collapsed()).toBe(true)
    expect(localStorage.getItem('t.ribbon')).toBe('1')
  })

  it('while collapsed no tab is selected and every tab offers Expand', () => {
    click('home')
    expect($('home').className).not.toContain('active')
    expect($('insert').className).not.toContain('active')
    expect($('home').dataset.tip).toBe($('insert').dataset.tip)
    expect($('home').dataset.tip).toMatch(/^Expand \(/)
  })

  it('pressing any tab while collapsed expands for good (no peek)', () => {
    click('home')
    click('insert')
    expect(collapsed()).toBe(false)
    expect($('insert').className).toContain('active')
    expect($('insert').dataset.tip).toMatch(/^Collapse \(/)
    expect($('home').dataset.tip).toBeUndefined()
    act(() => {
      document.body.dispatchEvent(new Event('pointerdown', { bubbles: true }))
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
      window.dispatchEvent(new Event('blur'))
    })
    expect(collapsed()).toBe(false)
    expect(localStorage.getItem('t.ribbon')).toBe('0')
  })

  // a browser double-click is click, click, dblclick
  const dblclick = (id: string) => {
    click(id)
    click(id)
    act(() => $(id).dispatchEvent(new MouseEvent('dblclick', { bubbles: true })))
  }

  it('double-clicking toggles relative to the state before the first click', () => {
    dblclick('insert')
    expect(collapsed()).toBe(true)
    dblclick('insert')
    expect(collapsed()).toBe(false)
    dblclick('home')
    expect(collapsed()).toBe(true)
    dblclick('home')
    expect(collapsed()).toBe(false)
  })

  it('a held shortcut (key repeat) does not re-toggle', () => {
    expect(
      isRibbonToggleShortcut(
        new KeyboardEvent('keydown', { key: 'F1', ctrlKey: true, repeat: true }),
      ),
    ).toBe(false)
  })

  it('Ctrl+F1 is the toggle shortcut, F1 alone is not', () => {
    expect(isRibbonToggleShortcut(new KeyboardEvent('keydown', { key: 'F1', ctrlKey: true }))).toBe(
      true,
    )
    expect(isRibbonToggleShortcut(new KeyboardEvent('keydown', { key: 'F1' }))).toBe(false)
    expect(
      isRibbonToggleShortcut(
        new KeyboardEvent('keydown', { key: 'F1', ctrlKey: true, altKey: true }),
      ),
    ).toBe(false)
  })

  it('reads the persisted flag, defaulting to expanded', () => {
    expect(readRibbonCollapsed('t.ribbon')).toBe(false)
    localStorage.setItem('t.ribbon', '1')
    expect(readRibbonCollapsed('t.ribbon')).toBe(true)
  })
})

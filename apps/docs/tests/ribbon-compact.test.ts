// Compact ribbon (#362): the middle density between "full" and "collapsed".
// ⌥⌘K / Ctrl+Alt+K cycles full ↔ compact; the band stays visible and the active
// tab stays highlighted, which is what separates compact from collapsed.
import { beforeEach, describe, expect, it } from 'vitest'
import { act, createElement, useState } from 'react'
import { createRoot } from 'react-dom/client'
import {
  isRibbonCompactShortcut,
  isRibbonToggleShortcut,
  readRibbonDensity,
  useRibbonCollapse,
  type RibbonDensity,
} from '@genoffice/ui'

const TABS = ['home', 'insert'] as const
// the opt-in labels: without `compact`/`expandFull` the hook stays two-state
const LABELS = {
  collapse: 'Collapse',
  expand: 'Expand',
  compact: 'Compact',
  expandFull: 'Expand the Ribbon',
}

function Ribbon() {
  const collapse = useRibbonCollapse('t.compact', LABELS)
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
const isCollapsed = () => rootEl().classList.contains('ribbon-collapsed')
const isCompact = () => rootEl().classList.contains('ribbon-compact')

/** ⌥⌘K on macOS, Ctrl+Alt+K elsewhere — the platform check is inside the hook. */
const pressCompact = () =>
  act(() => {
    const mac = navigator.platform.toLowerCase().includes('mac')
    window.dispatchEvent(
      new KeyboardEvent('keydown', {
        key: 'k',
        code: 'KeyK',
        metaKey: mac,
        ctrlKey: !mac,
        altKey: true,
      }),
    )
  })

/** the existing collapse shortcut, same platform split */
const pressCollapse = () =>
  act(() => {
    const mac = navigator.platform.toLowerCase().includes('mac')
    window.dispatchEvent(
      new KeyboardEvent('keydown', {
        key: mac ? 'r' : 'F1',
        metaKey: mac,
        ctrlKey: !mac,
        altKey: mac,
      }),
    )
  })

const mount = () =>
  act(() => {
    document.body.innerHTML = ''
    const host = document.createElement('div')
    document.body.append(host)
    createRoot(host).render(createElement(Ribbon))
  })

describe('ribbon compact density', () => {
  beforeEach(() => {
    localStorage.removeItem('t.compact')
    mount()
  })

  it('starts at full: no compact class, band visible, tab selected', () => {
    expect(isCompact()).toBe(false)
    expect(isCollapsed()).toBe(false)
    expect($('home').className).toContain('active')
  })

  it('the compact shortcut turns the band icon-only without collapsing it', () => {
    pressCompact()
    expect(isCompact()).toBe(true)
    expect(isCollapsed()).toBe(false)
    // the band element stays in the tree — only the density class changed
    expect(rootEl().querySelector('[data-ribbon-body]')).not.toBeNull()
  })

  it('compact keeps the active tab highlighted (collapsed does not)', () => {
    pressCompact()
    expect($('home').className).toContain('active')
    // the collapse shortcut is the other axis and does drop the selection
    pressCollapse()
    expect($('home').className).not.toContain('active')
  })

  it('the same shortcut goes back to full', () => {
    pressCompact()
    pressCompact()
    expect(isCompact()).toBe(false)
    expect(isCollapsed()).toBe(false)
  })

  it('collapse and compact are independent axes', () => {
    // from collapsed, ⌥⌘K reveals the band compactly rather than at full size
    pressCollapse()
    expect(isCollapsed()).toBe(true)
    pressCompact()
    expect(isCompact()).toBe(true)
    expect(isCollapsed()).toBe(false)
    // the band toggle from compact goes straight to collapsed, not to full
    pressCollapse()
    expect(isCollapsed()).toBe(true)
    pressCollapse()
    expect(isCollapsed()).toBe(false)
    expect(isCompact()).toBe(false)
  })

  it('persists each density under the shared storage key', () => {
    pressCompact()
    expect(localStorage.getItem('t.compact')).toBe('compact')
    pressCollapse()
    expect(localStorage.getItem('t.compact')).toBe('collapsed')
  })

  it('a legacy collapsed flag still reads back as collapsed', () => {
    localStorage.setItem('t.compact', '1')
    expect(readRibbonDensity('t.compact')).toBe('collapsed')
    localStorage.setItem('t.compact', '0')
    expect(readRibbonDensity('t.compact')).toBe('full')
    localStorage.setItem('t.compact', 'compact')
    expect(readRibbonDensity('t.compact')).toBe('compact')
  })

  it('remounts into the persisted density', () => {
    pressCompact()
    mount()
    expect(isCompact()).toBe(true)
  })

  it('the active tab advertises both ways out while compact', () => {
    const tip = $('home').dataset.tip
    expect(tip).toMatch(/Compact/)
    // the unselected tab stays quiet
    expect($('insert').dataset.tip).toBeUndefined()
  })

  it('a held compact key does not re-toggle', () => {
    const mac = navigator.platform.toLowerCase().includes('mac')
    act(() => {
      window.dispatchEvent(
        new KeyboardEvent('keydown', {
          key: 'k',
          code: 'KeyK',
          metaKey: mac,
          ctrlKey: !mac,
          altKey: true,
          repeat: true,
        }),
      )
    })
    expect(isCompact()).toBe(false)
  })
})

describe('isRibbonCompactShortcut', () => {
  it('matches only the alt-modified K chord', () => {
    const mac = navigator.platform.toLowerCase().includes('mac')
    const ev = (init: KeyboardEventInit) => new KeyboardEvent('keydown', init)
    expect(
      isRibbonCompactShortcut(ev({ code: 'KeyK', metaKey: mac, ctrlKey: !mac, altKey: true })),
    ).toBe(true)
    // bare K, K with the wrong modifier, and the collapse chord are all rejected
    expect(isRibbonCompactShortcut(ev({ code: 'KeyK' }))).toBe(false)
    expect(
      isRibbonCompactShortcut(ev({ code: 'KeyK', metaKey: mac, altKey: true, shiftKey: true })),
    ).toBe(false)
    expect(isRibbonCompactShortcut(ev({ code: 'KeyR', metaKey: mac, altKey: true }))).toBe(false)
    expect(
      isRibbonCompactShortcut(ev({ code: 'KeyK', metaKey: mac, altKey: true, repeat: true })),
    ).toBe(false)
  })

  it('does not collide with the collapse shortcut', () => {
    const mac = navigator.platform.toLowerCase().includes('mac')
    const collapseChord = new KeyboardEvent('keydown', {
      key: mac ? 'r' : 'F1',
      code: mac ? 'KeyR' : 'F1',
      metaKey: mac,
      ctrlKey: !mac,
      altKey: mac,
    })
    expect(isRibbonToggleShortcut(collapseChord)).toBe(true)
    expect(isRibbonCompactShortcut(collapseChord)).toBe(false)
  })
})

describe('two-state call sites are unaffected', () => {
  it('omitting the compact labels keeps the original boolean behaviour', () => {
    function Plain() {
      const collapse = useRibbonCollapse('t.plain', { collapse: 'Collapse', expand: 'Expand' })
      return createElement('div', { className: collapse.rootClass, id: 'plain' })
    }
    act(() => {
      document.body.innerHTML = ''
      const host = document.createElement('div')
      document.body.append(host)
      createRoot(host).render(createElement(Plain))
    })
    const el = document.getElementById('plain') as HTMLElement
    expect(el.className).toBe('ribbon-collapsible')
    expect(el.className).not.toContain('ribbon-compact')
    // the legacy key shape is preserved for apps that never opted in
    localStorage.removeItem('t.plain')
    expect(readRibbonDensity('t.plain')).toBe('full')
  })
})

// keep the type import referenced so the density union stays honest
export type { RibbonDensity }

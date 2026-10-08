/**
 * The document-level contextmenu decision (genoffice#1816): a grid right-click must
 * cancel the DOM event — Univer shows its own menu and never cancels it, so
 * a surviving event lets Electron stack the native menu on top. Editable
 * surfaces keep the native edit menu; the footer strip opens our stats menu.
 */
// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'

import { sheetContextMenuAction } from '../src/renderer/sheet-context-menu'

/** Builds `tag` nested under the given ancestors, outermost first.
 * `div|id=x` syntax sets attributes. */
function applyAttrs(node: Element, attrs: string[]): void {
  for (const attr of attrs) {
    const eq = attr.indexOf('=')
    node.setAttribute(eq === -1 ? attr : attr.slice(0, eq), eq === -1 ? '' : attr.slice(eq + 1))
  }
}

function el(tag: string, ...ancestors: string[]): Element {
  const [nodeTag = '', ...nodeAttrs] = tag.split('|')
  const node = document.createElement(nodeTag)
  applyAttrs(node, nodeAttrs)
  let current: Element = node
  for (const spec of [...ancestors].reverse()) {
    const [parentTag = '', ...attrs] = spec.split('|')
    const parent = document.createElement(parentTag)
    applyAttrs(parent, attrs)
    parent.appendChild(current)
    current = parent
  }
  return node
}

describe('sheet context menu action', () => {
  it('suppresses the native menu for a grid right-click so it cannot stack on Univer’s', () => {
    const cell = el('div', 'div|id=univer-container', 'div|class=univer-canvas')
    expect(sheetContextMenuAction(cell)).toEqual({
      statsMenu: false,
      suppressNativeMenu: true,
    })
  })

  it('keeps the native edit menu on editable surfaces (formula bar, in-cell editor)', () => {
    const input = el('input', 'div|id=univer-container', 'div|data-u-comp=formula-bar')
    expect(sheetContextMenuAction(input)).toEqual({
      statsMenu: false,
      suppressNativeMenu: false,
    })
    // a div inside the contenteditable editor host is editable too
    const insideEditor = el(
      'span',
      'div|id=univer-container',
      'div|class=univer-editor|contenteditable=true',
    )
    expect(sheetContextMenuAction(insideEditor)).toEqual({
      statsMenu: false,
      suppressNativeMenu: false,
    })
  })

  it('opens our stats menu for the footer strip below the sheet tabs', () => {
    // real footer shape: section[data-range-selector] > [tab strip, stats area]
    const section = document.createElement('section')
    section.setAttribute('data-range-selector', '')
    section.appendChild(document.createElement('div')) // the sheet tab strip
    const stats = document.createElement('div')
    section.appendChild(stats)
    const container = document.createElement('div')
    container.setAttribute('id', 'univer-container')
    container.appendChild(section)
    expect(sheetContextMenuAction(stats)).toEqual({
      statsMenu: true,
      suppressNativeMenu: true,
    })
  })

  it('leaves the sheet tab strip to Univer (suppresses the native menu, no stats menu)', () => {
    // the strip is the section's firstElementChild, so it is not ours to handle
    const tabs = el(
      'button',
      'div|id=univer-container',
      'section|data-range-selector=',
      'div|class=tab-strip',
    )
    expect(sheetContextMenuAction(tabs)).toEqual({
      statsMenu: false,
      suppressNativeMenu: true,
    })
  })

  it('ignores right-clicks outside the spreadsheet (ribbon, AI panel)', () => {
    const ribbon = el('button', 'header|class=excel-header')
    expect(sheetContextMenuAction(ribbon)).toEqual({
      statsMenu: false,
      suppressNativeMenu: false,
    })
    expect(sheetContextMenuAction(null)).toEqual({
      statsMenu: false,
      suppressNativeMenu: false,
    })
  })
})

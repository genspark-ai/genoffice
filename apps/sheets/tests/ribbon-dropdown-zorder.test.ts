// @vitest-environment jsdom
/**
 * Regression pin for the ribbon-header z-order lift (genoffice#1767).
 *
 * .excel-header is z-index: 10, so it opens a stacking context and a ribbon
 * dropdown popover's own z-index is trapped inside it, painting under the
 * grid's float layers (the slicer stack and the watch window, both 1000 in
 * the root order). The header is therefore lifted while one of ITS OWN
 * popovers is open.
 *
 * The defect pinned here is which popover counts. html.genoffice-popover-open
 * is toggled by every useDismissablePopover, including the Dropdown inside a
 * dialog; keying the lift on that global class lifted the whole header above
 * .dialog-backdrop (z-index: 40), un-dimming it and making it clickable over
 * an open dialog. "Dialogs always dismiss the popover first" only covers a
 * dialog opened FROM a popover, not a popover opened INSIDE a dialog. The lift
 * must key on the popover's HOST.
 *
 * jsdom applies no stylesheet and resolves no stacking context, so a
 * computed-style assertion here would be vacuous. What is assertable is the
 * real selector's match semantics, so the tests read the lift selector out of
 * styles.css and run it against DOM built by the real components (the ribbon's
 * ColorDropdown, SubtotalDialog's Dropdown) rather than against class names
 * transcribed by hand — the selector is checked against what the components
 * actually emit, and html.genoffice-popover-open is set by the real
 * useDismissablePopover rather than by the test.
 */
import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it } from 'vitest'

import { ColorDropdown } from '../src/renderer/ColorDropdown'
import { SubtotalDialog } from '../src/renderer/SubtotalDialog'

// jsdom rewrites import.meta.url to an http URL, so resolve from cwd instead.
const stylesheet = (): string => {
  const candidates = [
    resolve(process.cwd(), 'src/renderer/styles.css'),
    resolve(process.cwd(), 'apps/sheets/src/renderer/styles.css'),
  ]
  const found = candidates.find((path) => existsSync(path))
  if (!found) throw new Error(`styles.css not found; looked in ${candidates.join(', ')}`)
  return readFileSync(found, 'utf8')
}

const css = stylesheet()

/** .excel-header's own z-index; anything but this re-orders the header. */
const BASE_HEADER_Z_INDEX = 10

/**
 * Selector text of every rule that re-z-indexes .excel-header, i.e. the lift.
 * Read from the shipped stylesheet so the assertions below run against the
 * selector as committed rather than a copy of it.
 */
function liftSelectors(): string[] {
  return [...css.matchAll(/([^{}]*\.excel-header[^{}]*)\{([^{}]*)\}/g)]
    .map((m) => ({ selector: m[1]!.trim(), body: m[2]! }))
    .filter((rule) => /z-index\s*:/.test(rule.body))
    .filter((rule) => Number(/z-index\s*:\s*(-?\d+)/.exec(rule.body)![1]) !== BASE_HEADER_Z_INDEX)
    .map((rule) => rule.selector)
}

/** True when the committed CSS would lift this element's header. */
const isLifted = (element: Element): boolean => liftSelectors().some((s) => element.matches(s))

let root: Root | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  root = null
  document.documentElement.classList.remove('genoffice-popover-open')
  document.body.innerHTML = ''
})

async function render(element: React.ReactElement): Promise<void> {
  const container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(element)
    await Promise.resolve()
  })
}

const click = async (element: Element): Promise<void> => {
  await act(async () => {
    element.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }))
    await Promise.resolve()
  })
}

/** ExcelShell's shape: the ribbon header, then the sheet body hosting dialogs. */
function shell(ribbon: React.ReactNode, body: React.ReactNode = null): React.JSX.Element {
  return createElement(
    'main',
    { className: 'app-shell' },
    createElement('header', { className: 'excel-header' }, ribbon),
    createElement('div', { className: 'sheet-body' }, body),
  )
}

const header = (): Element => document.querySelector('.excel-header')!

describe('ribbon header z-order lift', () => {
  it('lifts the header for a popover hosted by the ribbon header', async () => {
    // The ribbon's color tools render ColorDropdown without `portal`, so the
    // palette lands inline in the header as .sheets-color-pop. Note this
    // component rolls its own dismissal and never joins the
    // useDismissablePopover refcount, so html.genoffice-popover-open is NOT set
    // here — keying the lift on that global class left the ribbon's colour
    // popovers trapped, which is the gap .sheets-color-pop in the selector
    // closes. MenuSelect/ShapeGallerySelect/EditableMenuSelect (the other
    // header hosts, .menu-select-drop) do set the class.
    await render(
      shell(
        createElement(ColorDropdown, {
          label: 'Fill color',
          value: '#ff0000',
          onPick: () => null,
        }),
      ),
    )
    await click(header().querySelector('button.color-well')!)

    expect(document.querySelector('.excel-header .sheets-color-pop')).not.toBeNull()
    expect(isLifted(header())).toBe(true)
  })

  it('does not lift the header for a popover hosted by a dialog', async () => {
    // SubtotalDialog renders .dialog-backdrop (z-index: 40) in the sheet body
    // and hosts the shared Dropdown; opening one of its dropdowns sets the very
    // same html.genoffice-popover-open class the old lift keyed on.
    await render(
      shell(
        null,
        createElement(SubtotalDialog, {
          fields: [
            { colIndex: 0, label: 'A' },
            { colIndex: 1, label: 'B' },
          ],
          onCreate: () => null,
          onClose: () => {},
        }),
      ),
    )
    const backdrop = document.querySelector('.dialog-backdrop')!
    await click(backdrop.querySelector('button.gs-dd-btn')!)

    const pop = document.querySelector('.gs-dd-pop')!
    expect(backdrop.contains(pop)).toBe(true)
    expect(header().contains(pop)).toBe(false)
    // Premise: the real useDismissablePopover inside the dialog is what puts
    // the global class on <html> — the condition the old rule matched on.
    expect(document.documentElement.classList.contains('genoffice-popover-open')).toBe(true)

    expect(isLifted(header())).toBe(false)
  })

  it('keeps a lift rule at all, and the base header z-index that makes it necessary', () => {
    expect(liftSelectors().length).toBeGreaterThan(0)
    const base = [...css.matchAll(/([^{}]*\.excel-header[^{}]*)\{([^{}]*)\}/g)].find((m) =>
      /z-index\s*:\s*10\s*;/.test(m[2]!),
    )
    expect(base).toBeDefined()
  })
})

// Insert > Icons dialog: Cancel, Escape and a click on the backdrop close it,
// except while the picked icons are being drawn and inserted, when none does.
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { IconsDialog } from '../src/renderer/components/IconsDialog'
import { loadIconLibrary } from '../src/renderer/icon-library'

const mounted: Array<{ root: Root; host: HTMLElement }> = []
afterEach(async () => {
  for (const { root, host } of mounted.splice(0)) {
    await act(async () => root.unmount())
    host.remove()
  }
})

async function mount() {
  await loadIconLibrary()
  const host = document.createElement('div')
  document.body.appendChild(host)
  const onClose = vi.fn()
  let finish = () => {}
  let fail = () => {}
  const onInsert = vi.fn(
    () =>
      new Promise<void>((resolve, reject) => {
        finish = resolve
        fail = () => reject(new Error('no bitmap'))
      }),
  )
  const root = createRoot(host)
  mounted.push({ root, host })
  await act(async () => {
    root.render(createElement(IconsDialog, { onInsert, onClose }))
  })
  const act1 = (run: () => void) => act(async () => run())
  const backdrop = host.querySelector<HTMLElement>('.modal-backdrop')!
  const button = (name: 'Cancel' | 'Insert') =>
    host.querySelector<HTMLButtonElement>(
      name === 'Insert' ? 'button.primary' : '.tco-actions-right button[type="button"]',
    )!
  const dismiss = {
    cancel: () => act1(() => button('Cancel').click()),
    escape: () =>
      act1(() => {
        host
          .querySelector('.icons-dialog-search')!
          .dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))
      }),
    backdrop: () =>
      act1(() => {
        backdrop.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }))
      }),
  }
  const pickAndInsert = async () => {
    await act1(() => host.querySelector<HTMLElement>('[data-icon="person"]')!.click())
    await act1(() => button('Insert').click())
    expect(onInsert).toHaveBeenCalledWith(['person'])
  }
  return {
    onClose,
    onInsert,
    dismiss,
    pickAndInsert,
    button,
    finish: () => act1(() => finish()),
    fail: () => act1(() => fail()),
  }
}

const PATHS = ['cancel', 'escape', 'backdrop'] as const

describe('dismissing the Icons dialog', () => {
  it.each(PATHS)('%s closes it while nothing is being inserted', async (path) => {
    const dialog = await mount()
    await dialog.dismiss[path]()
    expect(dialog.onClose).toHaveBeenCalledTimes(1)
    expect(dialog.onInsert).not.toHaveBeenCalled()
  })

  it.each(PATHS)('%s does nothing while the picked icons are being inserted', async (path) => {
    const dialog = await mount()
    await dialog.pickAndInsert()
    await dialog.dismiss[path]()
    expect(dialog.onClose).not.toHaveBeenCalled()
    await dialog.finish()
    expect(dialog.onInsert).toHaveBeenCalledTimes(1)
  })

  it.each(PATHS)('%s closes it again after an insert that failed', async (path) => {
    const dialog = await mount()
    await dialog.pickAndInsert()
    await dialog.fail()
    expect(dialog.button('Insert').disabled).toBe(false)
    await dialog.dismiss[path]()
    expect(dialog.onClose).toHaveBeenCalledTimes(1)
  })
})

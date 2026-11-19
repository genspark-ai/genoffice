// New from Template gallery: while one template is being confirmed or loaded,
// Create, a double-click, Choose a File and every way out do nothing.
import { afterEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { TemplateGalleryDialog } from '../src/renderer/components/MenuDialogs'

const mounted: Array<{ root: Root; host: HTMLElement }> = []
afterEach(async () => {
  for (const { root, host } of mounted.splice(0)) {
    await act(async () => root.unmount())
    host.remove()
  }
  vi.unstubAllGlobals()
})

async function mount() {
  vi.stubGlobal('desktop', {
    listTemplates: async () => [{ name: 'Letter.dotx' }, { name: 'Plain.dotx' }],
  })
  const host = document.createElement('div')
  document.body.appendChild(host)
  const settle: Array<(made: boolean) => void> = []
  const pending = () => new Promise<boolean>((resolve) => settle.push(resolve))
  const onCreate = vi.fn(pending)
  const onBrowse = vi.fn(pending)
  const onClose = vi.fn()
  const root = createRoot(host)
  mounted.push({ root, host })
  await act(async () => {
    root.render(createElement(TemplateGalleryDialog, { onCreate, onBrowse, onClose }))
  })
  const act1 = (run: () => void) => act(async () => run())
  const tile = (name: string) => host.querySelector<HTMLElement>(`[data-template="${name}"]`)!
  return {
    onCreate,
    onBrowse,
    onClose,
    form: host.querySelector<HTMLElement>('form')!,
    create: () => act1(() => host.querySelector<HTMLButtonElement>('button.primary')!.click()),
    browse: () =>
      act1(() => host.querySelector<HTMLButtonElement>('button[name="browse"]')!.click()),
    doubleClick: (name: string) =>
      act1(() => {
        tile(name).dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))
      }),
    cancel: () =>
      act1(() => host.querySelector<HTMLButtonElement>('.tco-actions-right button')!.click()),
    escape: () =>
      act1(() => {
        tile('Plain.dotx').dispatchEvent(
          new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }),
        )
      }),
    backdrop: () =>
      act1(() => {
        host
          .querySelector('.modal-backdrop')!
          .dispatchEvent(new MouseEvent('mousedown', { bubbles: true }))
      }),
    settle: (made: boolean) => act1(() => settle.shift()!(made)),
  }
}

describe('the template gallery while a template is being opened', () => {
  it('starts one new document however often it is activated', async () => {
    const dialog = await mount()
    await dialog.create()
    await dialog.create()
    await dialog.doubleClick('Plain.dotx')
    await dialog.browse()
    expect(dialog.onCreate).toHaveBeenCalledTimes(1)
    expect(dialog.onCreate).toHaveBeenCalledWith('Letter.dotx')
    expect(dialog.onBrowse).not.toHaveBeenCalled()
    expect(dialog.form.getAttribute('aria-busy')).toBe('true')
  })

  it('takes a second activation in the same tick as the first for one', async () => {
    const dialog = await mount()
    await act(async () => {
      const plain = document.querySelector<HTMLElement>('[data-template="Plain.dotx"]')!
      plain.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))
      plain.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))
    })
    expect(dialog.onCreate).toHaveBeenCalledTimes(1)
  })

  it('stays open until the pending one is settled', async () => {
    const dialog = await mount()
    await dialog.browse()
    await dialog.cancel()
    await dialog.escape()
    await dialog.backdrop()
    expect(dialog.onClose).not.toHaveBeenCalled()
  })

  it('works again after a cancelled picker or a declined replace', async () => {
    const dialog = await mount()
    await dialog.browse()
    await dialog.settle(false)
    expect(dialog.form.getAttribute('aria-busy')).toBe('false')
    await dialog.doubleClick('Plain.dotx')
    expect(dialog.onCreate).toHaveBeenCalledWith('Plain.dotx')
    await dialog.settle(false)
    await dialog.escape()
    expect(dialog.onClose).toHaveBeenCalledTimes(1)
  })
})

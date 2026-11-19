// Footnote and Endnote dialog: Insert and Apply answer to the Start at field
// that is on screen, never to the other kind's hidden one.
import { describe, expect, it, onTestFinished, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { FootnoteDialog } from '../src/renderer/components/FootnoteDialog'

type Props = Parameters<typeof FootnoteDialog>[0]

async function mount(settings: Props['settings'], multiSection = false) {
  const host = document.createElement('div')
  document.body.appendChild(host)
  const onApply = vi.fn<Props['onApply']>()
  const onInsert = vi.fn<Props['onInsert']>()
  let root: Root
  await act(async () => {
    root = createRoot(host)
    root.render(
      createElement(FootnoteDialog, {
        settings,
        multiSection,
        counts: { footnote: 0, endnote: 0 },
        onApply,
        onInsert,
        onConvert: () => undefined,
        onClose: () => undefined,
      }),
    )
  })
  onTestFinished(() => {
    act(() => root.unmount())
    host.remove()
  })
  const q = <T extends Element>(selector: string) => host.querySelector<T>(selector)!
  const type = async (input: HTMLInputElement, value: string) => {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
    await act(async () => {
      setter.call(input, value)
      input.dispatchEvent(new Event('input', { bubbles: true }))
    })
  }
  const click = async (el: Element) => {
    await act(async () => {
      ;(el as HTMLElement).click()
    })
  }
  return { q, type, click, onApply, onInsert }
}

describe('FootnoteDialog', () => {
  it('a cleared Start at blocks Insert only while its kind is the one shown', async () => {
    const { q, type, click, onApply, onInsert } = await mount({ footnote: {}, endnote: {} })
    const insert = q<HTMLButtonElement>('button.primary')
    await type(q<HTMLInputElement>('input[name="note-start"]'), '')
    expect(insert.disabled).toBe(true)
    await click(q('input[value="endnote"]'))
    expect(insert.disabled).toBe(false)
    await click(insert)
    expect(onInsert).toHaveBeenCalledWith('endnote', '')
    expect(onApply).not.toHaveBeenCalled()
    // back on footnotes the field shows the number that is in force again
    await click(q('input[value="footnote"]'))
    expect(q<HTMLInputElement>('input[name="note-start"]').value).toBe('1')
    expect(insert.disabled).toBe(false)
  })

  it('a document whose hidden kind starts at 0 can still insert and apply', async () => {
    const { q, click, onApply } = await mount({ footnote: {}, endnote: { numStart: 0 } })
    expect(q<HTMLButtonElement>('button.primary').disabled).toBe(false)
    const format = q<HTMLSelectElement>('select[name="note-format"]')
    await act(async () => {
      format.value = 'upperRoman'
      format.dispatchEvent(new Event('change', { bubbles: true }))
    })
    const apply = q<HTMLButtonElement>('button[name="apply"]')
    expect(apply.disabled).toBe(false)
    await click(apply)
    expect(onApply).toHaveBeenCalledWith(
      { footnote: { numFmt: 'upperRoman' }, endnote: { numStart: 0 } },
      'document',
    )
    expect(
      q<HTMLOptionElement>('select[name="note-apply-to"] option[value="section"]').disabled,
    ).toBe(true)
  })

  it('This section applies the numbering and leaves the placement to the document', async () => {
    const { q, click, onApply } = await mount(
      { footnote: { pos: 'beneathText', numStart: 2 }, endnote: {} },
      true,
    )
    const change = async (name: string, value: string) => {
      const select = q<HTMLSelectElement>(`select[name="${name}"]`)
      await act(async () => {
        select.value = value
        select.dispatchEvent(new Event('change', { bubbles: true }))
      })
    }
    await change('footnote-pos', 'pageBottom')
    await change('note-restart', 'eachSect')
    await change('note-apply-to', 'section')
    const pos = q<HTMLSelectElement>('select[name="footnote-pos"]')
    expect(pos.disabled).toBe(true)
    expect(pos.value).toBe('beneathText')
    await click(q<HTMLButtonElement>('button[name="apply"]'))
    expect(onApply).toHaveBeenCalledWith(
      { footnote: { numStart: 2, numRestart: 'eachSect' }, endnote: {} },
      'section',
    )
  })
})

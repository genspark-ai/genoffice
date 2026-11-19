import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { MutableRefObject } from 'react'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { StylePanel } from '../src/renderer/components/StylePanel'
import type { ComputedSnapshot } from '../src/renderer/preview/inspector-protocol'

const computed: ComputedSnapshot = {
  color: '#000000',
  fontSize: '16px',
  fontWeight: '400',
  fontStyle: 'normal',
  textAlign: 'left',
  background: '',
  backgroundImage: 'none',
  width: '100px',
  height: '20px',
  borderRadius: '0px',
  padding: '0px',
  opacity: '1',
  transform: '',
  marginLeft: '',
  marginRight: '',
  objectFit: 'fill',
}

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
})

function panelProps(
  draftRef: MutableRefObject<(() => void) | null>,
  onCustomCss: (css: string) => void,
) {
  return {
    tag: 'div',
    computed,
    textRun: null,
    onStyle: vi.fn(),
    onText: vi.fn(),
    onAttr: vi.fn(),
    onCustomCss,
    draftRef,
    pending: false,
    onRevert: vi.fn(),
    onClose: vi.fn(),
  }
}

async function renderPanel(
  draftRef: MutableRefObject<(() => void) | null>,
  onCustomCss: (css: string) => void,
) {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(createElement(StylePanel, panelProps(draftRef, onCustomCss)))
    await Promise.resolve()
  })
  return container
}

async function setCustomCss(input: HTMLInputElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set
  setter?.call(input, value)
  await act(async () => {
    input.dispatchEvent(new Event('input', { bubbles: true }))
    await Promise.resolve()
  })
}

describe('StylePanel custom CSS drafts', () => {
  it('commits through the host draft flusher', async () => {
    const draftRef: MutableRefObject<(() => void) | null> = { current: null }
    const onCustomCss = vi.fn()
    const host = await renderPanel(draftRef, onCustomCss)
    const input = host.querySelector<HTMLInputElement>('input[placeholder^="letter-spacing:"]')!
    await setCustomCss(input, 'letter-spacing: 2px')
    await act(async () => {
      draftRef.current?.()
      await Promise.resolve()
    })
    expect(onCustomCss).toHaveBeenCalledWith('letter-spacing: 2px')
  })

  it('commits when the panel unmounts', async () => {
    const draftRef: MutableRefObject<(() => void) | null> = { current: null }
    const onCustomCss = vi.fn()
    const host = await renderPanel(draftRef, onCustomCss)
    const input = host.querySelector<HTMLInputElement>('input[placeholder^="letter-spacing:"]')!
    await setCustomCss(input, 'box-shadow: 0 1px 2px red')
    await act(async () => {
      root!.unmount()
      await Promise.resolve()
    })
    root = null
    expect(onCustomCss).toHaveBeenCalledWith('box-shadow: 0 1px 2px red')
  })
})

describe('StylePanel text alignment', () => {
  it('offers justify and shows it as the selected value for a justified element', async () => {
    const draftRef: MutableRefObject<(() => void) | null> = { current: null }
    const onCustomCss = vi.fn()
    ;(computed as { textAlign: string }).textAlign = 'justify'
    const host = await renderPanel(draftRef, onCustomCss)
    const select = [...host.querySelectorAll('select')].find((s) =>
      [...s.options].some((o) => o.value === 'justify'),
    )!
    expect(select).toBeTruthy()
    const justifyOption = [...select.options].find((o) => o.value === 'justify')!
    // label text follows the UI locale; only presence and value are invariant
    expect(justifyOption.textContent).toBeTruthy()
    // a justified element reports justify, not a fallback to left
    expect(select.value).toBe('justify')
    ;(computed as { textAlign: string }).textAlign = 'left'
  })
})

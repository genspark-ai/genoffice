// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'

import { NotePictureHoverCard } from '../src/renderer/NotePicturesPane'
import type { NotePopupCard } from '../src/renderer/note-hover-card'

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

const cardOf = (over: Partial<NotePopupCard['picture']> = {}): NotePopupCard => ({
  sheetId: 'sheet-1',
  row: 3,
  column: 2,
  picture: {
    sheetId: 'sheet-1',
    sheetName: 'Data',
    row: 3,
    column: 2,
    author: 'Alice',
    text: '老照片扫描件',
    dataUrl: 'data:image/png;base64,AAA',
    ...over,
  },
})

describe('NotePictureHoverCard', () => {
  it('shows the author, the picture and the text under the popup; remove hits the cell', async () => {
    const onRemove = vi.fn()
    await mount(
      createElement(NotePictureHoverCard, {
        card: cardOf(),
        sessionId: 'session-1',
        onRemove,
      }),
    )
    const card = document.querySelector('.note-hover-card')
    expect(card).not.toBeNull()
    expect(card!.querySelector('.np-author')?.textContent).toBe('Alice')
    expect(card!.querySelector('.np-image')?.getAttribute('src')).toBe('data:image/png;base64,AAA')
    expect(card!.querySelector('.np-text')?.textContent).toBe('老照片扫描件')
    await act(async () => {
      card!.querySelector<HTMLButtonElement>('.np-remove')!.click()
    })
    expect(onRemove).toHaveBeenCalledWith('sheet-1', 3, 2)
  })

  it('renders nothing without a card, and no picture without bytes', async () => {
    await mount(
      createElement(NotePictureHoverCard, { card: null, sessionId: undefined, onRemove: vi.fn() }),
    )
    expect(document.querySelector('.note-hover-card')).toBeNull()

    await mount(
      createElement(NotePictureHoverCard, {
        card: {
          sheetId: 'sheet-1',
          row: 3,
          column: 2,
          picture: {
            sheetId: 'sheet-1',
            sheetName: 'Data',
            row: 3,
            column: 2,
            author: 'Alice',
            text: '老照片扫描件',
          },
        },
        sessionId: undefined,
        onRemove: vi.fn(),
      }),
    )
    expect(document.querySelector('.note-hover-card .np-image')).toBeNull()
    // the text still shows: the popup itself carries only the note text
    expect(document.querySelector('.note-hover-card .np-text')?.textContent).toBe('老照片扫描件')
  })

  it('stays unpositioned when Univer has no popup box to anchor under', async () => {
    await mount(
      createElement(NotePictureHoverCard, {
        card: cardOf(),
        sessionId: undefined,
        onRemove: vi.fn(),
      }),
    )
    const card = document.querySelector<HTMLElement>('.note-hover-card')!
    expect(card.style.left).toBe('')
    expect(card.style.top).toBe('')
  })

  it('anchors under the newest cell popup box', async () => {
    const popup = document.createElement('div')
    popup.className = 'cell-popup'
    document.body.appendChild(popup)
    vi.spyOn(popup, 'getBoundingClientRect').mockReturnValue({
      left: 40,
      top: 60,
      right: 300,
      bottom: 180,
      width: 260,
      height: 120,
      x: 40,
      y: 60,
      toJSON: () => ({}),
    } as DOMRect)
    await mount(
      createElement(NotePictureHoverCard, {
        card: cardOf(),
        sessionId: undefined,
        onRemove: vi.fn(),
      }),
    )
    const card = document.querySelector<HTMLElement>('.note-hover-card')!
    expect(card.style.left).toBe('40px')
    expect(card.style.top).toBe('186px')
  })
})

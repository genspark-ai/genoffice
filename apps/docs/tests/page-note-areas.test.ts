import { describe, expect, it, vi } from 'vitest'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { NoteInfo } from '@genoffice/docx-engine'
import { PageEndnotes, PageFootnotes } from '../src/renderer/components/PageNoteAreas'
import { endnotesAnchorY } from '../src/renderer/pagination'

function render(element: React.ReactElement): { container: HTMLElement; unmount: () => void } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  return {
    container,
    unmount: () => {
      act(() => root.unmount())
      container.remove()
    },
  }
}

const note = (id: string, text: string): NoteInfo => ({ id, text })

describe('PageEndnotes — canvas endnote area (Word parity)', () => {
  it('renders no heading, roman superscript markers, and the entry text', () => {
    const { container, unmount } = render(
      createElement(PageEndnotes, {
        notes: [note('1', 'first'), note('2', 'second')],
        top: null,
        onEdit: () => {},
        onDelete: () => {},
      }),
    )
    // Word shows no "Endnotes" heading: the area starts at the numbered entries
    expect(container.querySelector('.page-notes-label')).toBeNull()
    const sups = Array.from(container.querySelectorAll('.page-note sup'), (s) => s.textContent)
    expect(sups).toEqual(['i', 'ii'])
    expect(container.textContent).toContain('first')
    unmount()
  })

  it('anchors at the measured flow end when a top is provided', () => {
    const { container, unmount } = render(
      createElement(PageEndnotes, {
        notes: [note('1', 'x')],
        top: 421.5,
        onEdit: () => {},
        onDelete: () => {},
      }),
    )
    const area = container.querySelector('.page-endnotes') as HTMLElement
    expect(area.classList.contains('page-endnotes-anchored')).toBe(true)
    expect(area.style.top).toBe('421.5px')
    unmount()
  })

  it('falls back to flow stacking before the first measure (no anchor class)', () => {
    const { container, unmount } = render(
      createElement(PageEndnotes, {
        notes: [note('1', 'x')],
        top: null,
        onEdit: () => {},
        onDelete: () => {},
      }),
    )
    const area = container.querySelector('.page-endnotes') as HTMLElement
    expect(area.classList.contains('page-endnotes-anchored')).toBe(false)
    expect(area.style.top).toBe('')
    unmount()
  })

  it('forwards edit/delete to the handlers', () => {
    const onEdit = vi.fn()
    const onDelete = vi.fn()
    const { container, unmount } = render(
      createElement(PageEndnotes, { notes: [note('7', 'x')], top: 10, onEdit, onDelete }),
    )
    const [edit, del] = Array.from(container.querySelectorAll<HTMLButtonElement>('.page-note-btn'))
    act(() => edit.click())
    act(() => del.click())
    expect(onEdit).toHaveBeenCalledWith('7')
    expect(onDelete).toHaveBeenCalledWith('7')
    unmount()
  })
})

describe('PageFootnotes — end-of-document footnote list', () => {
  it('skips entries already shown in page gaps but keeps full-list numbering', () => {
    const { container, unmount } = render(
      createElement(PageFootnotes, {
        notes: [note('a', 'one'), note('b', 'two'), note('c', 'three')],
        skipIds: new Set(['a', 'b']),
        onEdit: () => {},
        onDelete: () => {},
      }),
    )
    const sups = Array.from(container.querySelectorAll('.page-note sup'), (s) => s.textContent)
    expect(sups).toEqual(['3'])
    expect(container.textContent).toContain('three')
    unmount()
  })

  it('renders nothing when every entry lives in a page gap', () => {
    const { container, unmount } = render(
      createElement(PageFootnotes, {
        notes: [note('a', 'one')],
        skipIds: new Set(['a']),
        onEdit: () => {},
        onDelete: () => {},
      }),
    )
    expect(container.querySelector('.page-notes')).toBeNull()
    unmount()
  })
})

describe('endnotesAnchorY — flow-end anchor for the endnote area', () => {
  const child = (opts: { cls?: string; bottom: number; height: number }): HTMLElement => {
    const el = document.createElement('div')
    if (opts.cls) el.className = opts.cls
    el.getBoundingClientRect = () =>
      ({ bottom: opts.bottom, height: opts.height, top: opts.bottom - opts.height }) as DOMRect
    return el
  }

  it('anchors below the last visible block, in layout px', () => {
    const pm = document.createElement('div')
    pm.append(child({ bottom: 300, height: 20 }), child({ bottom: 500, height: 30 }))
    expect(endnotesAnchorY(pm, 100, 2)).toBe(200)
  })

  it('skips page gaps, float hosts and zero-height markers at the tail', () => {
    const pm = document.createElement('div')
    pm.append(
      child({ bottom: 500, height: 30 }),
      child({ bottom: 501, height: 0 }),
      child({ cls: 'page-gap', bottom: 900, height: 200 }),
      child({ cls: 'page-float-host', bottom: 910, height: 10 }),
    )
    expect(endnotesAnchorY(pm, 0, 1)).toBe(500)
  })

  it('returns null when no visible block exists', () => {
    const pm = document.createElement('div')
    pm.append(child({ bottom: 10, height: 0 }))
    expect(endnotesAnchorY(pm, 0, 1)).toBeNull()
  })
})

describe('note-area CSS contract (document data, not chrome)', () => {
  const css = readFileSync(join(__dirname, '../src/renderer/styles.css'), 'utf8')
  const ruleOf = (selector: string): string => {
    const m = css.match(new RegExp(`${selector.replace(/[.[\]]/g, '\\$&')}\\s*\\{([^}]*)\\}`))
    expect(m, `rule ${selector} missing`).toBeTruthy()
    return m![1]
  }

  it('note-area markers inherit the area ink (no accent/link blue) and superscript', () => {
    const rule = ruleOf('.page-note sup')
    expect(rule).not.toContain('color')
    // flex rows ignore vertical-align: the raise comes from a relative offset
    expect(rule).toContain('position: relative')
    expect(rule).toMatch(/top:\s*-0\.\d+em/)
  })

  it('anchored endnote area is absolutely positioned (top set inline per measure)', () => {
    const rule = ruleOf('.page-endnotes-anchored')
    expect(rule).toContain('position: absolute')
  })

  it('the endnote heading style is gone with the heading itself', () => {
    expect(css).not.toContain('.page-notes-label')
  })
})

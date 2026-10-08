/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { Markdown, type MarkdownNav } from '../src/Markdown'
import { createFileNav, fileNavHref } from '../src/file-nav'

let host: HTMLDivElement
let root: ReturnType<typeof createRoot>

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  vi.unstubAllGlobals()
})

const render = (text: string, navs?: readonly MarkdownNav[], nav?: MarkdownNav): void => {
  act(() => {
    root.render(createElement(Markdown, { text, navs, nav }))
  })
}

const navLink = (): HTMLAnchorElement | null => host.querySelector('a.ai-md-nav')

const click = (el: HTMLAnchorElement): void =>
  act(() => {
    el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }))
  })

describe('Markdown multi-scheme navs', () => {
  it('renders a filenav link as a clickable nav anchor and reports the href', () => {
    const onNavigate = vi.fn()
    render(`see [report.docx](${fileNavHref('/Users/me/report.docx')})`, [
      createFileNav(onNavigate),
    ])
    const link = navLink()
    expect(link).not.toBeNull()
    expect(link!.textContent).toBe('report.docx')
    expect(link!.getAttribute('href')).toBe('filenav:///Users/me/report.docx')
    click(link!)
    expect(onNavigate).toHaveBeenCalledWith('/Users/me/report.docx')
  })

  it('matches each nav by its own scheme in one answer', () => {
    const gotoBlock = vi.fn()
    const openFile = vi.fn()
    render('[第3段](docnav://block/3) cites [report.docx](filenav:///Users/me/report.docx)', [
      createFileNav(openFile),
      { scheme: 'docnav://', onNavigate: gotoBlock },
    ])
    const links = Array.from(host.querySelectorAll('a.ai-md-nav'))
    expect(links).toHaveLength(2)
    click(links[0]!)
    expect(gotoBlock).toHaveBeenCalledWith('docnav://block/3')
    expect(openFile).not.toHaveBeenCalled()
    click(links[1]!)
    expect(openFile).toHaveBeenCalledWith('/Users/me/report.docx')
    expect(gotoBlock).toHaveBeenCalledTimes(1)
  })

  it('keeps non-matching schemes as dead literal text', () => {
    render('see [docs](https://example.com/a.md) and [page](pdfnav://page/2)', [
      createFileNav(vi.fn()),
    ])
    expect(navLink()).toBeNull()
    expect(host.textContent).toContain('[docs](https://example.com/a.md)')
    expect(host.textContent).toContain('[page](pdfnav://page/2)')
  })

  it('still honors the legacy single nav prop on its own', () => {
    const onNavigate = vi.fn()
    render('[go](docnav://block/7)', undefined, { scheme: 'docnav://', onNavigate })
    click(navLink()!)
    expect(onNavigate).toHaveBeenCalledWith('docnav://block/7')
  })

  it('accepts the legacy nav prop alongside a navs list', () => {
    const onNavigate = vi.fn()
    const openFile = vi.fn()
    render('[go](docnav://block/7) then [file](filenav:///a/b.md)', [createFileNav(openFile)], {
      scheme: 'docnav://',
      onNavigate,
    })
    const links = Array.from(host.querySelectorAll('a.ai-md-nav'))
    expect(links).toHaveLength(2)
    click(links[0]!)
    expect(onNavigate).toHaveBeenCalledWith('docnav://block/7')
    click(links[1]!)
    expect(openFile).toHaveBeenCalledWith('/a/b.md')
  })
})

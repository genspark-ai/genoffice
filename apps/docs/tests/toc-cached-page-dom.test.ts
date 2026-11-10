/**
 * TOC page display on open: the file's cached field results stay until the
 * document is edited; only lines without a stored result take the live page.
 */
import { describe, expect, it } from 'vitest'
import { backfillTocPageDom } from '../src/renderer/editor/toc-refresh'

function tocLine(anchor: string, stored: string): HTMLElement {
  const el = document.createElement('p')
  el.className = 'doc-toc-line'
  el.setAttribute('data-toc-anchor', anchor)
  const page = document.createElement('span')
  page.className = 'doc-toc-page'
  page.textContent = stored
  el.append(page)
  return el
}

const pageText = (lines: HTMLElement[]) =>
  lines.map((l) => l.querySelector('.doc-toc-page')!.textContent)

describe('backfillTocPageDom', () => {
  const live: Record<string, string> = { _Toc1: '17', _Toc2: '18', _Toc3: '19' }

  it('keeps the stored results of a pristine document', () => {
    const lines = [tocLine('_Toc1', '16'), tocLine('_Toc2', ''), tocLine('_Toc3', '20')]
    const n = backfillTocPageDom(lines, (a) => live[a], false)
    expect(n).toBe(1)
    expect(pageText(lines)).toEqual(['16', '18', '20'])
    // a later clean pass with a settled layout moves only the line we filled
    expect(backfillTocPageDom(lines, (a) => ({ ...live, _Toc2: '19' })[a], false)).toBe(1)
    expect(pageText(lines)).toEqual(['16', '19', '20'])
  })

  it('takes the live layout once the document is dirty', () => {
    const lines = [tocLine('_Toc1', '16'), tocLine('_Toc2', ''), tocLine('_Toc9', '20')]
    const n = backfillTocPageDom(lines, (a) => live[a], true)
    expect(n).toBe(2)
    // an unmeasured anchor keeps whatever it shows
    expect(pageText(lines)).toEqual(['17', '18', '20'])
  })
})

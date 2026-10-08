import { describe, expect, it } from 'vitest'
import type { SectionInfo } from '@genoffice/docx-engine'
import { assignSections } from '../src/renderer/pagination-sections'
import type { BlockBox } from '../src/renderer/pagination-types'

const sec = (first: number, last: number, startType: SectionInfo['startType']): SectionInfo => ({
  settings: {
    pageWidth: 11906,
    pageHeight: 16838,
    orientation: 'portrait',
    marginTop: 1440,
    marginRight: 1440,
    marginBottom: 1440,
    marginLeft: 1440,
    pageBorder: false,
    columns: 1,
  },
  startType,
  firstBlockIndex: first,
  lastBlockIndex: last,
  sectPrXml: '',
  titlePg: false,
  headerRefs: {},
  footerRefs: {},
})

function block(docxIndex: number, top: number, height: number, marginTopPx: number): BlockBox {
  const el = document.createElement('p')
  el.style.marginTop = `${marginTopPx}px`
  document.body.appendChild(el)
  return { docxIndex, top, height, el } as BlockBox
}

// measureBlocks folds the inter-block margin into the previous block's space-after
function flow(startType: SectionInfo['startType']) {
  const a = block(0, 0, 20, 0)
  const b = block(1, 36, 20, 16)
  a.spaceAfterPx = 16
  a.height += 16
  assignSections([a, b], [sec(0, 0, 'nextPage'), sec(1, 1, startType)])
  return { a, b }
}

describe('space-before of the first paragraph of a forced section start', () => {
  it('stays with the paragraph (Word keeps it on the section page)', () => {
    const { a, b } = flow('nextPage')
    expect(a.spaceAfterPx).toBe(0)
    expect(a.height).toBe(20)
    expect(b.top).toBe(20)
    expect(b.height).toBe(36)
    expect(b.spaceBeforePx).toBe(16)
    expect(b.leadFoldPx).toBe(16)
  })

  it('a continuous section keeps the folded space-after (no page starts there)', () => {
    const { a, b } = flow('continuous')
    expect(a.spaceAfterPx).toBe(16)
    expect(b.top).toBe(36)
    expect(b.spaceBeforePx).toBeUndefined()
  })
})

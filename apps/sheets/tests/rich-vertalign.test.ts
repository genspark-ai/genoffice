import { BaselineOffset } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { toRichTextDocument } from '../src/renderer/univer-sync'
import type { WorkbookRichRun } from '../src/shared/desktop-api'

const run = (text: string, extra: Partial<WorkbookRichRun> = {}): WorkbookRichRun => ({
  text,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  ...extra,
})

describe('rich-run vertAlign', () => {
  it('maps subscript/superscript runs to Univer baseline offsets', () => {
    const p = toRichTextDocument('Cr2O3 g/cm3', [
      run('Cr'),
      run('2', { vertAlign: 'subscript', size: 14 }),
      run('O'),
      run('3', { vertAlign: 'subscript', size: 14 }),
      run(' g/cm'),
      run('3', { vertAlign: 'superscript' }),
    ])
    const ts = p?.body?.textRuns?.map((t) => t.ts)
    expect(ts?.[1]?.va).toBe(BaselineOffset.SUBSCRIPT)
    expect(ts?.[1]?.fs).toBe(14)
    expect(ts?.[3]?.va).toBe(BaselineOffset.SUBSCRIPT)
    expect(ts?.[5]?.va).toBe(BaselineOffset.SUPERSCRIPT)
    expect(ts?.[0]?.va).toBeUndefined()
  })

  it('treats a vertAlign-only run as formatted (does not collapse to the base)', () => {
    const base = { fs: 11 }
    const p = toRichTextDocument('x2', [run('x'), run('2', { vertAlign: 'subscript' })], base)
    expect(p?.body?.textRuns?.[1]?.ts?.va).toBe(BaselineOffset.SUBSCRIPT)
    expect(p?.body?.textRuns?.[1]?.ts?.fs).toBe(11)
  })
})

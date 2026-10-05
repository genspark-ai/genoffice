import { describe, expect, it, vi } from 'vitest'

// file-actions pulls in the offscreen render stack (Konva/canvas) at import time;
// only the file-name helpers are under test here.
vi.mock('../src/renderer/export-render', () => ({ renderSlidesToPngBase64: vi.fn() }))
vi.mock('../src/renderer/export-pages', () => ({ renderSlidesToPdfPages: vi.fn() }))
vi.mock('../src/renderer/i18n/locale', () => ({
  t: (key: string) => key,
}))

import { baseName } from '../src/shared/base-name'
import { exportBaseName } from '../src/renderer/file-actions'
import type { ActionCtx } from '../src/renderer/action-context'

/** Only `path` is read; the rest of the ctx is irrelevant to the file name. */
function ctxWithPath(path: string | undefined): ActionCtx {
  return { path } as unknown as ActionCtx
}

describe('baseName', () => {
  it('returns the file name of a posix path', () => {
    expect(baseName('/Users/me/Documents/deck.pptx')).toBe('deck.pptx')
  })

  it('returns the file name of a Windows path', () => {
    // node's basename only strips the host platform's separator, so splitting on
    // '/' alone handed the whole path back and every export name became invalid
    expect(baseName('C:\\Users\\me\\deck.pptx')).toBe('deck.pptx')
  })

  it('handles mixed separators, a bare file name and an empty path', () => {
    expect(baseName('C:\\Users\\me\\sub\\deck.pptx')).toBe('deck.pptx')
    expect(baseName('deck.pptx')).toBe('deck.pptx')
    expect(baseName('')).toBe('')
  })
})

describe('exportBaseName', () => {
  it('strips the .pptx extension from a Windows session path', () => {
    // op.baseName reaches main as join(dir, `${baseName}-01.png`)
    expect(exportBaseName(ctxWithPath('C:\\Users\\me\\deck.pptx'))).toBe('deck')
  })

  it('strips the .pptx extension from a posix session path', () => {
    expect(exportBaseName(ctxWithPath('/Users/me/Documents/deck.pptx'))).toBe('deck')
  })

  it('falls back to the untitled label when no file is open', () => {
    expect(exportBaseName(ctxWithPath(undefined))).toBe('appUntitledPresentation')
  })
})

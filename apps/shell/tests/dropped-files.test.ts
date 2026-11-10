import { beforeEach, describe, expect, it, vi } from 'vitest'

import { handleDroppedFiles } from '../src/main/dropped-files'
import type { DroppedFilesDeps } from '../src/main/dropped-files'

function fakeDeps(overrides: Partial<DroppedFilesDeps> = {}): DroppedFilesDeps & {
  opened: string[]
  revealed: ReturnType<typeof vi.fn>
  warned: string[]
} {
  const opened: string[] = []
  const revealed = vi.fn()
  const warned: string[] = []
  return {
    opened,
    revealed,
    warned,
    openDocumentPath: vi.fn((path: string) => {
      opened.push(path)
      return true
    }),
    revealShellWindow: revealed,
    showWarning: vi.fn((message: string) => {
      warned.push(message)
    }),
    unsupportedMessage: (exts) => `unsupported: ${exts.join(', ')}`,
    ...overrides,
  }
}

describe('handleDroppedFiles', () => {
  beforeEach(() => vi.clearAllMocks())

  it('opens every supported path and reveals the shell', () => {
    const deps = fakeDeps()
    handleDroppedFiles(['/tmp/a.docx', '/tmp/b.xlsx'], deps)
    expect(deps.opened).toEqual(['/tmp/a.docx', '/tmp/b.xlsx'])
    expect(deps.revealed).toHaveBeenCalledOnce()
    expect(deps.warned).toEqual([])
  })

  it('still reveals the shell when a path is already open in a tab', () => {
    const deps = fakeDeps({
      // routeDocumentPath activates the existing tab and returns true either way
      openDocumentPath: vi.fn(() => true),
    })
    handleDroppedFiles(['/tmp/a.docx'], deps)
    expect(deps.revealed).toHaveBeenCalledOnce()
  })

  it('warns once with the combined extensions for known-unsupported drops', () => {
    const deps = fakeDeps()
    handleDroppedFiles(['/tmp/old.doc', '/tmp/deck.pages', '/tmp/note.rtf'], deps)
    expect(deps.opened).toEqual([])
    expect(deps.revealed).not.toHaveBeenCalled()
    expect(deps.warned).toEqual(['unsupported: doc, pages, rtf'])
  })

  it('opens the supported files and warns about the rest in a mixed drop', () => {
    const deps = fakeDeps()
    handleDroppedFiles(['/tmp/new.docx', '/tmp/old.doc'], deps)
    expect(deps.opened).toEqual(['/tmp/new.docx'])
    expect(deps.revealed).toHaveBeenCalledOnce()
    expect(deps.warned).toEqual(['unsupported: doc'])
  })

  it('ignores junk payloads and unrecognized file types entirely', () => {
    const deps = fakeDeps()
    handleDroppedFiles(null, deps)
    handleDroppedFiles(['/tmp/photo.png', 42, ''], deps)
    expect(deps.opened).toEqual([])
    expect(deps.revealed).not.toHaveBeenCalled()
    expect(deps.warned).toEqual([])
  })
})

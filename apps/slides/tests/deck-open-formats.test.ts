import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  OPENABLE_DECK_EXTENSIONS,
  deckExtension,
  isOpenableDeckPath,
} from '../src/shared/deck-open-formats'

/**
 * The open guard and the open dialog's picker filter read this one list. An
 * `.otp` template missing from it is unreachable from the standalone app even
 * though the shell offers it.
 */
describe('OPENABLE_DECK_EXTENSIONS', () => {
  it('covers every extension the guard accepts', () => {
    expect(OPENABLE_DECK_EXTENSIONS).toEqual(['pptx', 'ppt', 'odp', 'otp'])
    for (const extension of OPENABLE_DECK_EXTENSIONS) {
      expect(isOpenableDeckPath(`/tmp/Quarterly Review.${extension.toUpperCase()}`)).toBe(true)
    }
  })

  it('rejects the neighbouring formats Slides does not open', () => {
    for (const path of ['/tmp/notes.txt', '/tmp/keynote.key', '/tmp/book.ods']) {
      expect(isOpenableDeckPath(path)).toBe(false)
    }
  })
})

describe('deckExtension', () => {
  it('lowercases the trailing extension and reports none when absent', () => {
    expect(deckExtension('/tmp/Quarterly Review.OTP')).toBe('otp')
    expect(deckExtension('/tmp/no-extension')).toBe('')
  })

  it('never turns a dotfile into an import', () => {
    // the trailing segment of a dotfile is not a format, so the import route
    // leaves such a path alone whatever the helper makes of it
    expect(isOpenableDeckPath('/tmp/.hidden')).toBe(false)
  })
})

/**
 * The picker and the copy live in the Electron main module, which no unit test
 * can import without standing up the whole editor. Asserting that the import
 * route reads the shared list and writes atomically is what keeps the three
 * copies of this route — the guard, the dialog, the write — from drifting apart.
 */
describe('the ODF import route in the Slides main module', () => {
  const source = readFileSync(join(import.meta.dirname, '../src/main/slides-main.ts'), 'utf8')

  it('offers exactly the extensions the guard accepts', () => {
    expect(source).toContain('extensions: [...OPENABLE_DECK_EXTENSIONS]')
    expect(source).not.toMatch(/extensions: \[[^\]]*'odp'[^\]]*\]/)
    expect(source).toContain('const extension = deckExtension(filePath)')
  })

  it('writes the converted copy atomically, as the Docs and shell routes do', () => {
    // a plain writeFile leaves a half-written .pptx where the user expects a
    // deck when the process dies mid-write
    expect(source).toContain('await atomicWriteFile(r.filePath, Buffer.from(result.pptx))')
    expect(source).toContain(
      "import { atomicWriteFile } from '@genoffice/electron-utils/atomic-write'",
    )
  })
})

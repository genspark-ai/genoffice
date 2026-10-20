import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { runGuardedCandidate, runGuardedDocumentAction } from '../src/renderer/doc-dirty'
import {
  OPENABLE_DOC_EXTENSIONS,
  findOpenableDocPath,
  isOpenableDocPath,
} from '../src/shared/open-file'

describe('findOpenableDocPath', () => {
  it('finds Finder and Explorer document arguments case-insensitively', () => {
    expect(
      findOpenableDocPath(['/Applications/GenOffice Docs.app', '/tmp/Quarterly Plan.docx']),
    ).toBe('/tmp/Quarterly Plan.docx')
    expect(findOpenableDocPath(['GenOffice Docs.exe', 'C:\\Users\\Me\\REPORT.DOCX'])).toBe(
      'C:\\Users\\Me\\REPORT.DOCX',
    )
    // an .odt argument routes through the import filter like the shell open
    expect(findOpenableDocPath(['GenOffice Docs', '/tmp/notes.odt'])).toBe('/tmp/notes.odt')
    // and so does the .ott template, in either case
    expect(findOpenableDocPath(['GenOffice Docs', '/tmp/letter.ott'])).toBe('/tmp/letter.ott')
    expect(findOpenableDocPath(['GenOffice Docs', '/tmp/LETTER.OTT'])).toBe('/tmp/LETTER.OTT')
  })

  it('ignores Electron switches and unrelated files', () => {
    expect(
      findOpenableDocPath(['GenOffice Docs', '--inspect=document.docx', '/tmp/notes.txt']),
    ).toBeNull()
  })
})

/**
 * The argv route, the external-open guard and the open dialog's picker filter
 * all read this one list. A template left out of it is unreachable from the
 * standalone app while the shell still offers it in its own dialog.
 */
describe('OPENABLE_DOC_EXTENSIONS', () => {
  it('covers every extension the argv route and the guard accept', () => {
    expect(OPENABLE_DOC_EXTENSIONS).toEqual(['docx', 'odt', 'ott'])
    for (const extension of OPENABLE_DOC_EXTENSIONS) {
      expect(isOpenableDocPath(`/tmp/Quarterly Plan.${extension.toUpperCase()}`)).toBe(true)
    }
  })

  it('rejects the neighbouring formats Docs does not open', () => {
    for (const path of ['/tmp/notes.txt', '/tmp/legacy.doc', '/tmp/deck.odp', '/tmp/book.ods']) {
      expect(isOpenableDocPath(path)).toBe(false)
    }
  })
})

/**
 * The external-open guard and the picker filter live in the Electron main
 * module, which no unit test can import. Asserting both sites reference the one
 * shared list is what keeps the dialog and the guard from drifting apart.
 */
describe('the open dialog and the external-open guard', () => {
  const source = readFileSync(join(import.meta.dirname, '../src/main/docs-main.ts'), 'utf8')

  it('offers exactly the extensions the guard accepts', () => {
    expect(source).toContain('extensions: [...OPENABLE_DOC_EXTENSIONS]')
    expect(source).not.toMatch(/extensions: \[[^\]]*'odt'[^\]]*\]/)
  })

  it('guards the external-open route with the same helper', () => {
    expect(source).toContain('if (!isOpenableDocPath(filePath)) return')
  })
})

describe('runGuardedCandidate', () => {
  it('chooses first and guards only once a candidate exists', async () => {
    const calls: string[] = []

    await expect(
      runGuardedCandidate(
        async () => {
          calls.push('guard')
          return true
        },
        async () => {
          calls.push('choose')
          return { path: '/tmp/candidate.docx' }
        },
        async (candidate) => {
          calls.push(`commit:${candidate.path}`)
        },
      ),
    ).resolves.toBe(true)

    expect(calls).toEqual(['choose', 'guard', 'commit:/tmp/candidate.docx'])
  })

  it('never runs the guard when the picker is cancelled', async () => {
    const confirm = vi.fn(async () => true)
    const commit = vi.fn()

    await expect(runGuardedCandidate(confirm, async () => null, commit)).resolves.toBe(false)
    expect(confirm).not.toHaveBeenCalled()
    expect(commit).not.toHaveBeenCalled()
  })

  it('does not commit a candidate when the replacement is rejected', async () => {
    const choose = vi.fn(async () => ({ path: '/tmp/candidate.docx' }))
    const commit = vi.fn()

    await expect(runGuardedCandidate(async () => false, choose, commit)).resolves.toBe(false)
    expect(choose).toHaveBeenCalledTimes(1)
    expect(commit).not.toHaveBeenCalled()
  })

  it('defers the guard for a candidate the caller marks as not a document yet', async () => {
    const confirm = vi.fn(async () => true)
    const commit = vi.fn()
    const encrypted = { needsPassword: true as const, path: '/tmp/locked.docx' }

    await expect(
      runGuardedCandidate(
        confirm,
        async () => encrypted,
        commit,
        (candidate) => !('needsPassword' in candidate),
      ),
    ).resolves.toBe(true)
    expect(confirm).not.toHaveBeenCalled()
    expect(commit).toHaveBeenCalledWith(encrypted)
  })
})

describe('runGuardedDocumentAction', () => {
  it('keeps the current document when the replacement guard is canceled', async () => {
    const replace = vi.fn()

    await expect(runGuardedDocumentAction(async () => false, replace)).resolves.toBe(false)
    expect(replace).not.toHaveBeenCalled()
  })

  it('runs the replacement only after the guard accepts', async () => {
    const calls: string[] = []

    await expect(
      runGuardedDocumentAction(
        async () => {
          calls.push('guard')
          return true
        },
        async () => {
          calls.push('replace')
        },
      ),
    ).resolves.toBe(true)
    expect(calls).toEqual(['guard', 'replace'])
  })
})

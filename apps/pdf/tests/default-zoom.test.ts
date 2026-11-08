// @vitest-environment node
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { readPdfDefaultZoom } from '../src/main/default-zoom'
import { normalizePdfDefaultZoom } from '../src/shared/default-zoom'
import { openingViewState } from '../src/renderer/view-state'
import type { PdfViewState } from '../src/renderer/view-state'

const saved: PdfViewState = { page: 3, frac: 0.25, scale: 1.25, fitMode: null }

describe('PDF opening zoom priority', () => {
  it.each(['width', 'page'] as const)(
    '%s overrides saved numeric zoom while retaining the reading position',
    (mode) => {
      const result = openingViewState(mode, saved)
      expect(result).toEqual({ ...saved, fitMode: mode })
      expect(saved.fitMode).toBeNull()
    },
  )

  it('lets fit page override a saved fit-width mode', () => {
    expect(openingViewState('page', { ...saved, fitMode: 'width' })).toEqual({
      ...saved,
      fitMode: 'page',
    })
  })

  it.each([null, 'page', 'width'] as const)(
    'last-used restores the per-file %s mode',
    (fitMode) => {
      const history = { ...saved, fitMode }
      expect(openingViewState('last-used', history)).toEqual(history)
    },
  )

  it.each([undefined, null, 'invalid', 2, {}, []])(
    'preserves legacy history for an absent or invalid preference: %j',
    (preference) => {
      expect(normalizePdfDefaultZoom(preference)).toBe('last-used')
      expect(openingViewState(preference, saved)).toEqual(saved)
    },
  )

  it.each(['last-used', undefined, 'width'])(
    'opens an unknown file at fit width for %s',
    (preference) => {
      expect(openingViewState(preference, null)).toEqual({
        page: 1,
        frac: 0,
        scale: 1,
        fitMode: 'width',
      })
    },
  )

  it('opens an unknown file at fit page when selected', () => {
    expect(openingViewState('page', null)).toEqual({ page: 1, frac: 0, scale: 1, fitMode: 'page' })
  })
})

const dirs: string[] = []
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true })
})

describe('PDF opening preference on disk', () => {
  it('reads the latest choice on each open and tolerates absent or corrupt settings', () => {
    const dir = mkdtempSync(join(tmpdir(), 'pdf-zoom-'))
    dirs.push(dir)
    const path = join(dir, 'app-settings.json')
    expect(readPdfDefaultZoom(path)).toBe('last-used')
    for (const value of ['width', 'page', 'last-used'] as const) {
      writeFileSync(path, JSON.stringify({ pdfDefaultZoom: value, language: 'en' }))
      expect(readPdfDefaultZoom(path)).toBe(value)
    }
    for (const invalid of ['{broken', 'null', '[]', '{}', '{"pdfDefaultZoom":false}']) {
      writeFileSync(path, invalid)
      expect(readPdfDefaultZoom(path)).toBe('last-used')
    }
  })
})

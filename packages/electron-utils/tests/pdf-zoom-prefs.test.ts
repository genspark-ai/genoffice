import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { readPdfDefaultZoom, writePdfDefaultZoom } from '../src/pdf-zoom-prefs'

const settingsFile = () => join(mkdtempSync(join(tmpdir(), 'pdf-zoom-')), 'app-settings.json')

describe('pdf default zoom preference', () => {
  it('defaults to fit width when unset, corrupt or unknown', () => {
    const p = settingsFile()
    expect(readPdfDefaultZoom(p)).toBe('width')
    writeFileSync(p, '{nope')
    expect(readPdfDefaultZoom(p)).toBe('width')
    writeFileSync(p, JSON.stringify({ pdfDefaultZoom: 'huge' }))
    expect(readPdfDefaultZoom(p)).toBe('width')
  })

  it('persists page/last and drops the key for the default, keeping other settings', () => {
    const p = settingsFile()
    writeFileSync(p, JSON.stringify({ language: 'en' }))
    writePdfDefaultZoom(p, 'page')
    expect(readPdfDefaultZoom(p)).toBe('page')
    writePdfDefaultZoom(p, 'last')
    expect(readPdfDefaultZoom(p)).toBe('last')
    writePdfDefaultZoom(p, 'width')
    expect(JSON.parse(readFileSync(p, 'utf8'))).toEqual({ language: 'en' })
  })
})

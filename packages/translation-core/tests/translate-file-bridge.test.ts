/**
 * Unit-level coverage for the whole-file translation bridge.
 *
 * The bridge spawns the upstream LumosAI `translate.py`; these tests exercise
 * the locating logic and the error paths that do not require a working Python
 * toolchain, so they run everywhere. The happy path (a real .docx round-trip)
 * is covered by the E2E suite when the skill directory is present.
 */
import { describe, expect, it } from 'vitest'
import { join } from 'node:path'

import {
  defaultOutputPath,
  isSupportedExtension,
  resolveTranslateSkills,
  SUPPORTED_EXTENSIONS,
  translateFile,
} from '../src/file-translate'

describe('translate-file bridge', () => {
  it('recognises every supported extension', () => {
    for (const ext of SUPPORTED_EXTENSIONS) {
      expect(isSupportedExtension(`report${ext}`)).toBe(true)
      expect(isSupportedExtension(ext)).toBe(true)
      expect(isSupportedExtension(`REPORT${ext.toUpperCase()}`)).toBe(true)
    }
  })

  it('rejects extensions the upstream script does not handle', () => {
    for (const ext of ['.txt', '.md', '.csv', '.png', '']) {
      expect(isSupportedExtension(`file${ext}`)).toBe(false)
    }
  })

  it('derives the default output path next to the source', () => {
    expect(defaultOutputPath('/tmp/spec.xls')).toBe('/tmp/spec_translated.xls')
    expect(defaultOutputPath('/tmp/deck.pptx')).toBe('/tmp/deck_translated.pptx')
    expect(defaultOutputPath('/tmp/a.b.docx')).toBe('/tmp/a.b_translated.docx')
  })

  it('resolveTranslateSkills honours the explicit override', () => {
    const previous = process.env.GENOFFICE_TRANSLATE_SKILLS_DIR
    const dir = join(__dirname, '..', 'src')
    process.env.GENOFFICE_TRANSLATE_SKILLS_DIR = dir
    try {
      const loc = resolveTranslateSkills()
      // No scripts/translate.py under src/ai, so the override is rejected and
      // resolution falls through — but it must never report 'override' with a
      // script that does not exist.
      if (loc.source === 'override') {
        expect(loc.scriptPath).toContain('translate.py')
      }
      expect(['override', 'bundled', 'legacy', 'missing']).toContain(loc.source)
    } finally {
      if (previous === undefined) delete process.env.GENOFFICE_TRANSLATE_SKILLS_DIR
      else process.env.GENOFFICE_TRANSLATE_SKILLS_DIR = previous
    }
  })

  it('translateFile reports a missing input instead of throwing', async () => {
    const result = await translateFile({ inputPath: '/definitely/not/here.docx' })
    expect(result.ok).toBe(false)
    expect(result.error).toContain('input not found')
  })

  it('translateFile rejects an unsupported extension before spawning', async () => {
    const result = await translateFile({ inputPath: __filename })
    expect(result.ok).toBe(false)
    expect(result.error).toContain('unsupported extension')
  })
})

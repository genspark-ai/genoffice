/**
 * Curated marketplace entries must never declare tool names the host does not
 * actually register with pi. The translation family is the most common source
 * of drift: previously the catalog advertised `detect_file_type`,
 * `translate_docx`, `translate_pdf`, `kb_p`, ... — none of which existed, so the
 * agent tried to call phantom tools and failed with "unknown tool".
 *
 * This guard pins the tools each curated translate entry may claim.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { searchMarketplace } from '../src/shell/skills'
import { createOfficeSession } from '@genoffice/agent-runtime'
import { createTranslateSkillExtension } from '@genoffice/agent-skills/extensions/translate-skill'
import { ALL_TRANSLATE_TOOL_NAMES } from '@genoffice/agent-skills'

const TMP_DATA = `/tmp/genoffice-marketplace-honesty-${Date.now()}`
process.env.DATA_DIR = TMP_DATA
process.env.GENOFFICE_DATA_DIR = TMP_DATA

describe('marketplace curated entries', () => {
  it('only declare tools that translate-skill actually registers', async () => {
    const session = await createOfficeSession({
      cwd: TMP_DATA,
      agentDir: TMP_DATA,
      extensionMode: 'print',
      extensionFactories: [createTranslateSkillExtension()],
    })
    try {
      const realTools = new Set(session.session.getAllTools().map((t) => t.name))
      const phantom = ['detect_file_type', 'translate_docx', 'translate_pdf', 'translate_pptx', 'translate_xls', 'translate_xlsx', 'extract_docx_text', 'extract_pptx_text', 'extract_sheet_text', 'kb_add', 'kb_import', 'kb_export', 'kb_validate']
      for (const name of phantom) {
        expect(realTools.has(name), `translate-skill must not advertise phantom tool: ${name}`).toBe(false)
      }
      for (const name of ALL_TRANSLATE_TOOL_NAMES) {
        expect(realTools.has(name), `translate-skill should advertise ${name}`).toBe(true)
      }
    } finally {
      session.dispose()
    }
  })

  it('translate-config, translate-docx, translate-pdf, translate-ppt, translate-xls and translate entries never claim phantom tools', () => {
    const results = searchMarketplace({ sort: 'name' })
    const ids = new Set([...results.skills, ...results.plugins].map((e) => e.id))
    for (const id of ['translate', 'translate-config', 'translate-docx', 'translate-pdf', 'translate-ppt', 'translate-xls']) {
      expect(ids.has(id), `expected curated entry ${id} in marketplace`).toBe(true)
      const entry = [...results.skills, ...results.plugins].find((e) => e.id === id)!
      const phantom = ['detect_file_type', 'translate_docx', 'translate_pdf', 'translate_pptx', 'translate_xls', 'translate_xlsx', 'kb_add', 'kb_import', 'kb_export', 'kb_validate']
      for (const name of phantom) {
        expect(entry.tools.includes(name), `${id} must not advertise phantom tool ${name}`).toBe(false)
      }
    }
  })
})

describe('skill handlers invalidate the live pi session', () => {
  it('export invalidateLivePiSession (smoke)', () => {
    const source = readFileSync(
      new URL('../src/shell/skills.ts', import.meta.url),
      'utf8',
    )
    for (const handler of [
      "registerHandle('home:install-skill',",
      "registerHandle('home:uninstall-skill',",
      "registerHandle('home:toggle-skill',",
      "registerHandle('home:marketplace-delete-upload',",
    ]) {
        const idx = source.indexOf(handler)
        expect(idx).toBeGreaterThan(-1)
        // The first statement inside the handler body must be the invalidation.
        const body = source.slice(idx, source.indexOf('  }', idx))
        expect(body).toMatch(/invalidateLivePiSession\(\)/)
      }
  })
})

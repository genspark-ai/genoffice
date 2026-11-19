import { describe, expect, it } from 'vitest'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { spawnSync } from 'node:child_process'
import { pathToFileURL } from 'node:url'
import { createBlankPptx } from '../src/index'
import { xmllintAvailable } from '../../../tools/ooxml-validate/validate-pptx.mjs'

// Exercise the real validator from a path with spaces even in a space-free CI checkout.
describe.skipIf(!xmllintAvailable() && !process.env.CI)('schema validator paths', () => {
  it('validates a clean deck when the schema directory contains spaces', async () => {
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'genoffice schema path '))
    try {
      const root = path.resolve(__dirname, '../../..')
      fs.cpSync(path.join(root, 'tools/ooxml-validate'), path.join(tmp, 'validator'), {
        recursive: true,
      })
      fs.symlinkSync(path.join(root, 'node_modules'), path.join(tmp, 'node_modules'), 'junction')
      const deck = path.join(tmp, 'clean.pptx')
      fs.writeFileSync(deck, await createBlankPptx())
      const validator = pathToFileURL(path.join(tmp, 'validator/validate-pptx.mjs')).href
      const result = spawnSync(
        process.execPath,
        [
          '--input-type=module',
          '-e',
          'const { validatePptx } = await import(process.argv[1]); console.log(JSON.stringify(await validatePptx(process.argv[2])))',
          validator,
          deck,
        ],
        { encoding: 'utf8' },
      )
      expect(result.error).toBeUndefined()
      expect(result.status, result.stderr).toBe(0)
      expect(JSON.parse(result.stdout)).toEqual([])
    } finally {
      fs.rmSync(tmp, { recursive: true, force: true })
    }
  })
})

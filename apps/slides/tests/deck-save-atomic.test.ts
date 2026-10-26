import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

/**
 * The deck saves — Ctrl+S, Save As, autosave and the MCP save — live in the
 * Electron main module, which no unit test can import without standing up the
 * whole editor. savePptxToFile streams into the path it is given and truncates
 * it first, so a deck saved onto its own path is cut off for as long as the
 * stream runs and stays cut off if the process dies (genoffice#2123). Every save has to
 * hand it the sibling temp that atomicWriteVia renames over the deck afterwards.
 */
describe('deck saves in the Slides main module', () => {
  const source = readFileSync(join(import.meta.dirname, '../src/main/slides-main.ts'), 'utf8')

  it('stream the package into the temp that atomicWriteVia publishes', () => {
    expect(source).toContain(
      "import { atomicWriteFile, atomicWriteVia } from '@genoffice/electron-utils/atomic-write'",
    )
    expect(source).toContain('atomicWriteVia(filePath, (tmp) => savePptxToFile(opened, tmp))')
  })

  it("never hand savePptxToFile the deck's own path", () => {
    // the helper's call, writing to the temp, is the only one allowed
    expect(source.match(/savePptxToFile\([^)]*\)/g)).toEqual(['savePptxToFile(opened, tmp)'])
    // Ctrl+S, Save As, autosave, MCP save
    expect(source.match(/await writeDeckFile\(session\.opened, /g)).toHaveLength(4)
  })
})

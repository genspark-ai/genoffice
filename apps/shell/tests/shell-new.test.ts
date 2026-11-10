import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { resolve } from 'node:path'
import { runInNewContext } from 'node:vm'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import { openPptx } from '@genoffice/pptx-engine'

const require = createRequire(import.meta.url)
// Inspect packaging metadata without downloading Electron or compiling OCR helpers.
// Template files themselves are read from disk below.
const configModule = { exports: {} as ReturnType<typeof require> }
runInNewContext(readFileSync(resolve(import.meta.dirname, '../electron-builder.cjs'), 'utf8'), {
  module: configModule,
  __dirname: resolve(import.meta.dirname, '..'),
  process: { platform: 'linux', arch: 'x64', env: {} },
  require: (id: string) =>
    id === 'node:fs' ? { ...require(id), existsSync: () => true } : require(id),
})
const config = configModule.exports

function template(ext: string): Buffer {
  const resource = config.win.extraResources.find(
    (entry: { to: string }) => entry.to === 'shell-new',
  )
  expect(resource, 'Windows package must ship the Explorer templates').toBeDefined()
  return readFileSync(resolve(import.meta.dirname, '..', resource.from, `blank.${ext}`))
}

describe('Windows Explorer New templates', () => {
  it('opens the packaged Word template as one empty paragraph', async () => {
    const doc = await parseDocx(template('docx'))
    const visible = doc.blocks.filter((block) => !block.hidden)
    expect(visible).toHaveLength(1)
    expect(visible[0].type).toBe('paragraph')
    expect(doc.styles.has('Heading1')).toBe(true)
  })

  it('ships an Excel workbook with one empty worksheet', async () => {
    const zip = await JSZip.loadAsync(template('xlsx'), { checkCRC32: true })
    const workbook = await zip.file('xl/workbook.xml')!.async('string')
    expect(workbook.match(/<sheet\s/g)).toHaveLength(1)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toMatch(/<sheetData\s*\/>|<sheetData>\s*<\/sheetData>/)
    expect(await zip.file('_rels/.rels')!.async('string')).toContain('xl/workbook.xml')
  })

  it('opens the packaged PowerPoint template as one empty slide', async () => {
    const opened = await openPptx(template('pptx'))
    expect(opened.deck.slides).toHaveLength(1)
    expect(opened.deck.slides[0].elements).toHaveLength(0)
  })

  it('provides nonempty Windows file type labels for the New menu', () => {
    for (const ext of ['docx', 'xlsx', 'pptx']) {
      const association = config.fileAssociations.find(
        (entry: { ext: string }) => entry.ext === ext,
      )
      expect(association.description).toBe(association.name)
    }
  })
})

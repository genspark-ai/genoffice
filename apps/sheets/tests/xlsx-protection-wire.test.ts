import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { XlsxSidecarClient } from '../src/main/xlsx-sidecar-client'
import { workbookRangeResultSchema } from '../src/shared/desktop-api'
import { buildCompatibilityFixture } from './fixture-builder'

it.each([
  {
    name: 'modern',
    attributes:
      'sheet="1" algorithmName="SHA-512" hashValue="aGFzaA==" saltValue="c2FsdA==" spinCount="100000"',
    password: {
      algorithmName: 'SHA-512',
      hashValue: 'aGFzaA==',
      saltValue: 'c2FsdA==',
      spinCount: 100000,
    },
  },
  { name: 'legacy', attributes: 'sheet="1" password="CBEB"', password: { legacy: 'CBEB' } },
  { name: 'no password', attributes: 'sheet="1"', password: undefined },
])(
  'validates the real Rust read-range response for $name sheet protection',
  async ({ attributes, password }) => {
    const directory = await mkdtemp(join(tmpdir(), 'xlsx-protection-wire-'))
    const path = join(directory, 'protected.xlsx')
    const zip = await JSZip.loadAsync(await buildCompatibilityFixture())
    const worksheet = (await zip.file('xl/worksheets/sheet1.xml')!.async('string')).replace(
      '</sheetData>',
      `</sheetData><sheetProtection ${attributes}/>`,
    )
    zip.file('xl/worksheets/sheet1.xml', worksheet)
    const bytes = await zip.generateAsync({ type: 'nodebuffer' })
    await writeFile(path, bytes)
    const client = new XlsxSidecarClient(
      fileURLToPath(new URL('../native/xlsx-engine/target/release/xlsx-sidecar', import.meta.url)),
    )
    let sessionId: string | null = null
    try {
      const opened = (await client.open(path)) as { sessionId: string; sheets: { id: string }[] }
      sessionId = opened.sessionId
      let raw: any
      for (let i = 0; i < 200; i++) {
        raw = await client.readRange({
          sessionId,
          sheetId: opened.sheets[0]!.id,
          range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
        })
        if (raw.indexingComplete) break
        await new Promise((resolve) => setTimeout(resolve, 5))
      }
      expect(raw.indexingComplete).toBe(true)
      // Validate the actual subprocess JSON at the same schema as workbook:read-range.
      const result = workbookRangeResultSchema.parse(raw)
      expect(result.sheetProtection).toMatchObject({
        protected: true,
        hasPassword: password !== undefined,
      })
      expect(result.sheetProtection?.password).toEqual(password)
      expect(result.cells.map((cell) => cell.value)).toEqual(['Old', 10])
      expect(await readFile(path)).toEqual(bytes)
    } finally {
      if (sessionId) await client.close(sessionId)
      client.stop()
      await rm(directory, { recursive: true, force: true })
    }
  },
)

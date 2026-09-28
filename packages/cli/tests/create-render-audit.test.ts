/**
 * create --audit / --render: --audit runs the local geometry audit on the bytes
 * just written; --render prints the deck through the app's headless export
 * (faked here), so those assertions skip on Windows like render.test.ts does.
 */
import { existsSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { imageBlocks } from '../src/mcp/run'
import { fakeApp, run, tempDir, writeMinimalPdf } from './helpers'

function deckSpec(dir: string): string {
  const spec = join(dir, 'deck.json')
  writeFileSync(
    spec,
    JSON.stringify({
      pages: [
        {
          background: '#101A2B',
          elements: [
            {
              type: 'text',
              x: 80,
              y: 60,
              w: 600,
              h: 60,
              paragraphs: [
                { runs: [{ text: 'Hello deck', sizePt: 32, bold: true, color: '#FFFFFF' }] },
              ],
            },
          ],
        },
      ],
    }),
  )
  return spec
}

describe('create --audit / --render', () => {
  it('rejects --render/--audit for other types instead of ignoring them', async () => {
    const dir = tempDir()
    const md = join(dir, 'a.md')
    writeFileSync(md, '# x')
    const r = await run([
      'create',
      '--type',
      'docx',
      '--from',
      md,
      '--out',
      join(dir, 'a.docx'),
      '--render',
      '--json',
    ])
    expect(r.code).toBe(1)
    expect(r.json().message).toContain('--type pptx')
  })

  it('MCP image blocks pick up detail.previews, not just detail.files', () => {
    const dir = tempDir()
    const png = join(dir, 'preview.png')
    writeFileSync(png, Buffer.from([0x89, 0x50, 0x4e, 0x47]))
    const { images, omitted } = imageBlocks({
      status: 'ok',
      command: 'create',
      summary: 'created deck',
      // spec-directory builds list page file names in detail.files; only the
      // object entries with a path (detail.previews) are PNGs to attach
      detail: { files: ['01.json', '02.json'], previews: [{ slide: 0, path: png }] },
    })
    expect(omitted).toBe(0)
    expect(images).toHaveLength(1)
    expect(images[0]!.mimeType).toBe('image/png')
  })

  it('--audit reports the geometry audit of the deck it just built', async () => {
    const dir = tempDir()
    const out = join(dir, 'deck.pptx')
    const r = await run([
      'create',
      '--type',
      'pptx',
      '--spec',
      deckSpec(dir),
      '--out',
      out,
      '--audit',
      '--json',
    ])
    expect(r.code).toBe(0)
    const audit = r.json().detail.audit
    expect(audit.counts).toEqual({ error: 0, warning: 0 })
    expect(audit.issues).toEqual([])
    expect(audit.slides).toEqual([expect.objectContaining({ slide: 0 })])
  })

  it('--render writes previews into the named directory', async () => {
    if (process.platform === 'win32') return
    const dir = tempDir()
    const out = join(dir, 'deck.pptx')
    const shots = join(dir, 'previews')
    const app = fakeApp(dir, writeMinimalPdf(join(dir, 'page.pdf')))
    const r = await run(
      [
        'create',
        '--type',
        'pptx',
        '--spec',
        deckSpec(dir),
        '--out',
        out,
        '--render',
        shots,
        '--json',
      ],
      { env: { ...process.env, GENOFFICE_APP_BIN: app } },
    )
    expect(r.code).toBe(0)
    const previews = r.json().detail.previews as {
      slide: number
      path: string
      width: number
      height: number
    }[]
    expect(previews).toEqual([
      { slide: 0, path: join(shots, 'deck-01.png'), width: 612, height: 792 },
    ])
    expect(existsSync(previews[0]!.path)).toBe(true)
  })

  it('a bare --render defaults to <out>-previews beside the deck', async () => {
    if (process.platform === 'win32') return
    const dir = tempDir()
    const out = join(dir, 'deck.pptx')
    const app = fakeApp(dir, writeMinimalPdf(join(dir, 'page.pdf')))
    const r = await run(
      ['create', '--type', 'pptx', '--spec', deckSpec(dir), '--out', out, '--render', '--json'],
      { env: { ...process.env, GENOFFICE_APP_BIN: app } },
    )
    expect(r.code).toBe(0)
    const previews = r.json().detail.previews as { path: string }[]
    expect(previews[0]!.path).toBe(join(dir, 'deck-previews', 'deck-01.png'))
  })
})

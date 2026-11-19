import { existsSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { run, tempDir } from './helpers'

// the in-repo real .ppt the ppt2pptx package's own regression test uses
const PPT_FIXTURE = join(__dirname, '../../file-parse/tests/fixtures/legacy-sample.ppt')

describe('genoffice convert legacy .ppt → .pptx', () => {
  it('converts the real fixture deck to a package that reopens as a pptx', async () => {
    expect(existsSync(PPT_FIXTURE), 'legacy-sample.ppt fixture exists').toBe(true)
    const dir = tempDir()
    const out = join(dir, 'deck.pptx')
    const result = await run(['convert', PPT_FIXTURE, '--to', 'pptx', '--out', out, '--json'])
    expect(result.code).toBe(0)
    expect(result.json().detail.slides).toBeGreaterThan(0)
    expect(Array.isArray(result.json().detail.warnings)).toBe(true)

    // the output must be a real pptx package with slides
    const zip = await JSZip.loadAsync(
      await import('node:fs').then((fs) => fs.promises.readFile(out)),
    )
    expect(zip.file('[Content_Types].xml')).not.toBeNull()
    expect(
      Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length,
    ).toBe(result.json().detail.slides)

    // a second conversion of the same input is deterministic
    const again = join(dir, 'again.pptx')
    const second = await run(['convert', PPT_FIXTURE, '--to', 'pptx', '--out', again, '--json'])
    expect(second.code).toBe(0)
    expect(second.json().detail.slides).toBe(result.json().detail.slides)
  })

  it('reports a damaged .ppt through the stable error model, not a crash', async () => {
    // a CFB signature but no valid structure inside: the container layer's
    // own checks answer with a stable LegacyPptError code in detail
    const bytes = new Uint8Array(512 + 4)
    const view = new DataView(bytes.buffer)
    bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0)
    view.setUint16(24, 0x0009, true) // sector shift 512
    const dir = tempDir()
    const input = join(dir, 'damaged.ppt')
    const { writeFileSync } = await import('node:fs')
    writeFileSync(input, bytes)
    const result = await run(['convert', input, '--to', 'pptx', '--json'])
    expect(result.code).not.toBe(0)
    const envelope = JSON.parse(result.stdout) as {
      detail?: { reason?: string }
    }
    expect(typeof envelope.detail?.reason).toBe('string')
    expect(envelope.detail?.reason!.length).toBeGreaterThan(0)
  })
})

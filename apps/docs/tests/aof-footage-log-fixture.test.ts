import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { parseDocx } from '@genoffice/docx-engine'
import { describe, expect, it } from 'vitest'
import { safeExternalUrl } from '@genoffice/electron-utils'
import { DOCS_LINK_PROTOCOLS } from '../src/shared/link-protocols'

const FIXTURE = join(__dirname, '..', 'fixtures', 'aof-footage-log-example.docx')

describe('aof-footage-log example fixture', () => {
  it('parses with aof-review:// hyperlinks that survive the click-path allowlist', async () => {
    const parsed = await parseDocx(await readFile(FIXTURE))
    const hrefs: string[] = []
    for (const block of parsed.blocks) {
      for (const run of block.runs ?? []) {
        const href = (run as { link?: { href?: string } }).link?.href
        if (href) hrefs.push(href)
      }
    }
    const aof = hrefs.filter((h) => h.startsWith('aof-review://'))
    expect(aof.length).toBeGreaterThanOrEqual(1)
    // Every log link is a seek with a timecode the phase-5 grammar accepts.
    for (const h of aof) {
      const url = new URL(h)
      expect(['seek', 'open']).toContain(url.host)
      if (url.host === 'seek') {
        expect(url.searchParams.get('file')).toBeTruthy()
        expect(url.searchParams.get('tc')).toMatch(/^\d{1,2}:[0-5]\d:[0-5]\d(:\d{1,2})?$/)
      }
      // ...and passes the exact gate the click path uses.
      expect(safeExternalUrl(h, { allowedProtocols: DOCS_LINK_PROTOCOLS })).toBe(h)
    }
  })
})

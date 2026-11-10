import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import { describe, expect, it } from 'vitest'

interface LicenseTextModule {
  licenseText(name: string, dir: string): string | null
}

describe('third-party license aggregation', () => {
  it('includes the PDFium wrapper and binary redistribution terms', async () => {
    const root = join(__dirname, '../../..')
    const modulePath = pathToFileURL(join(root, 'tools/license-text.mjs')).href
    const { licenseText } = (await import(modulePath)) as LicenseTextModule
    const text = licenseText('@embedpdf/pdfium', join(root, 'node_modules/@embedpdf/pdfium'))

    expect(text).toContain('MIT License')
    expect(text).toContain('Copyright 2014 PDFium Authors')
    expect(text).toContain('Apache License')
  })

  it('generates the notice consumed by every shell packaging entrypoint', () => {
    const root = join(__dirname, '../../..')
    execFileSync(process.execPath, [join(root, 'tools/gen-third-party-notices.mjs')], {
      cwd: root,
      stdio: 'pipe',
    })
    const text = readFileSync(join(root, 'apps/shell/build/THIRD-PARTY-NOTICES.txt'), 'utf8')
    expect(text).toContain('@embedpdf/pdfium')
    expect(text).toContain('Copyright 2014 PDFium Authors')
    expect(text).toContain('Apache License')

    const shellPackage = JSON.parse(
      readFileSync(join(root, 'apps/shell/package.json'), 'utf8'),
    ) as { scripts: Record<string, string> }
    for (const target of ['dist:mac', 'dist:win', 'dist:linux']) {
      expect(shellPackage.scripts[target]).toContain('npm run notices')
    }
    const builder = readFileSync(join(root, 'apps/shell/electron-builder.cjs'), 'utf8')
    expect(builder).toContain('ensureThirdPartyNotices()')
  })
})

import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const here = dirname(fileURLToPath(import.meta.url))
const read = (rel: string) => readFileSync(join(here, '..', rel), 'utf8')

const appSrc = read('src/renderer/App.tsx')
const shellSrc = read('src/renderer/ExcelShell.tsx')
const iconSrc = read('src/renderer/ribbon-icons.tsx')

describe('Save As quick-access command', () => {
  it('renders a dedicated, independently enabled button beside Save', () => {
    expect(shellSrc).toMatch(
      /<SaveIcon \/>[\s\S]*?disabled=\{!canSaveAs\}[\s\S]*?onClick=\{onSaveAs\}[\s\S]*?<SaveAsIcon \/>/,
    )
  })

  it('is available for clean file-backed workbooks and invokes save-as mode', () => {
    expect(appSrc).toContain('canSaveAs={workbookFile !== null}')
    expect(appSrc).toContain("onSaveAs={() => void handleSave('save-as')}")
  })

  it('uses a distinct floppy-and-pencil icon', () => {
    expect(iconSrc).toContain('export function SaveAsIcon()')
    expect(iconSrc).toContain('pencil overlay')
  })
})

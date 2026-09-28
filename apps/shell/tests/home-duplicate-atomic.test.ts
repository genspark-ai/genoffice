import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

const source = readFileSync(join(__dirname, '../src/main/index.ts'), 'utf8')

function section(text: string, start: string, end: string): string {
  return text.slice(text.indexOf(start), text.indexOf(end, text.indexOf(start)))
}

const newSheetTab = section(source, 'async function newSheetTab', 'function surfaceNewTabError')
const newPdfTab = section(source, 'async function newPdfTab', 'function startQueuedWorkbookNudge')
const duplicateFile = section(source, 'HOME_CHANNELS.duplicateFile', 'HOME_CHANNELS.deleteFiles')
const mcpBlankSheet = section(
  source,
  'async function openBlankSheetsTabForMcp',
  'const tabId = tabManager',
)

describe('Home file creation is atomic', () => {
  it('publishes a blank spreadsheet through a temporary file', () => {
    expect(newSheetTab).toContain('await atomicWriteFile(backingPath, await blankXlsxBuffer())')
    expect(newSheetTab).toContain('await atomicWriteFile(suggestedPath, await blankXlsxBuffer())')
    expect(newSheetTab).not.toMatch(/writeFileSync\(/)
  })

  it('publishes the MCP blank spreadsheet through a temporary file', () => {
    expect(mcpBlankSheet).toContain('await atomicWriteFile(filePath, await blankXlsxBuffer())')
    expect(mcpBlankSheet).not.toMatch(/writeFileSync\(/)
  })

  it('publishes a blank PDF through a temporary file', () => {
    expect(newPdfTab).toContain('await atomicWriteFile(filePath, await blankPdfBuffer())')
    expect(newPdfTab).not.toMatch(/writeFileSync\(filePath/)
  })

  it('publishes a duplicate through a temporary file', () => {
    expect(duplicateFile).toContain('await atomicCopyFile(path, target)')
    expect(duplicateFile).not.toMatch(/copyFileSync\(path, target\)/)
    // a kernel copy, never the whole file through renderer-visible memory
    expect(duplicateFile).not.toMatch(/readFileSync\(path\)/)
  })
})

describe('Home duplicate error handling', () => {
  it('localizes a failed duplicate instead of letting the raw error escape', () => {
    expect(duplicateFile).toContain('showErrorDialog(shellWindow, tm(')
    expect(duplicateFile).toMatch(/catch \(err\) \{[\s\S]*?return/)
  })

  it('does not record a recent for a duplicate that was never written', () => {
    const write = duplicateFile.indexOf('await atomicCopyFile(path, target)')
    const catchBlock = duplicateFile.indexOf('catch (err)')
    const record = duplicateFile.indexOf('recordRecentFile(target)')
    expect(catchBlock).toBeGreaterThan(write)
    expect(record).toBeGreaterThan(catchBlock)
  })

  it('still returns void to the renderer', () => {
    expect(duplicateFile).toMatch(/duplicateFile, async \(_event, path: unknown\)/)
  })
})

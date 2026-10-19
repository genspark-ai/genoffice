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
const sharedDuplicate = readFileSync(join(__dirname, '../src/main/tab-duplicate.ts'), 'utf8')
const duplicateWrapper = section(source, 'function duplicateFile(', 'function tabMenuActions')
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
    expect(duplicateFile).toContain('await duplicateFile(path)')
    expect(sharedDuplicate).toContain('await atomicCopyFile(path, target)')
    expect(sharedDuplicate).not.toMatch(/copyFileSync\(path, target\)/)
    // a kernel copy, never the whole file through renderer-visible memory
    expect(sharedDuplicate).not.toMatch(/readFileSync\(path\)/)
  })
})

describe('Home duplicate error handling', () => {
  it('localizes a failed duplicate instead of letting the raw error escape', () => {
    expect(duplicateWrapper).toContain(
      "showErrorDialog(shellWindow, tm('errDuplicateFailed'), error)",
    )
    expect(sharedDuplicate).toMatch(/catch \(error\) \{[\s\S]*?deps.onError\(error\)[\s\S]*?return/)
  })

  it('does not record a recent for a duplicate that was never written', () => {
    const write = sharedDuplicate.indexOf('await atomicCopyFile(path, target)')
    const catchBlock = sharedDuplicate.indexOf('catch (error)')
    const record = sharedDuplicate.indexOf('deps.recordRecentFile(target)')
    expect(catchBlock).toBeGreaterThan(write)
    expect(record).toBeGreaterThan(catchBlock)
  })

  it('still returns void to the renderer', () => {
    expect(duplicateFile).toMatch(/duplicateFile, async \(_event, path: unknown\)/)
  })
})

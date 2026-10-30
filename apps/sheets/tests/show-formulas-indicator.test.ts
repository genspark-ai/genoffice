/**
 * Formula view (sheetView/@showFormulas) is a per-sheet state that a save
 * persists, so a sheet can open showing formulas instead of values — a table
 * that reads as broken until you know to press Ctrl+` again.
 *
 * The one-off toast that announced the switch is long gone by the time anyone
 * notices, so the pane has to keep saying which sheet is in that mode.
 */
import { readFileSync, readdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { formulaViewSheets } from '../src/renderer/formula-view'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

const here = dirname(fileURLToPath(import.meta.url))
const read = (rel: string) => readFileSync(join(here, '..', rel), 'utf8')

const shellSrc = read('src/renderer/ExcelShell.tsx')
const appSrc = read('src/renderer/App.tsx')
const css = read('src/renderer/styles.css')

describe('formula view indicator', () => {
  it('the flag follows the sheet, not the workbook', () => {
    const state = {
      showFormulaSheets: new Set(['sheet-2']),
    } as unknown as LazyWorkbookState
    expect(formulaViewSheets(state).has('sheet-2')).toBe(true)
    expect(formulaViewSheets(state).has('sheet-1')).toBe(false)
  })

  it('the status bar carries the flag while the sheet shows formulas', () => {
    expect(shellSrc).toMatch(/\{showFormulas && \(\s*<span className="status-flag"/)
    expect(shellSrc).toMatch(/t\('appShowingFormulasActive'\)/)
  })

  it('the flag is styled, not a bare word in the message line', () => {
    expect(css).toMatch(/\.status-flag \{[^}]*font-weight: 600;/s)
    expect(css).toMatch(/\.status-flag \{[^}]*background: var\(--sheets-chip-bg\);/s)
  })

  it('the ribbon button reads as pressed while the mode is on', () => {
    expect(shellSrc).toMatch(/active=\{showFormulas\}[\s\S]{0,200}?appShowFormulas/)
  })

  it('the app tells the shell which sheet is in formula view', () => {
    expect(appSrc).toMatch(/showFormulas=\{activeShowFormulas\}/)
    expect(appSrc).toMatch(/formulaViewSheets\(state\)\.has\(worksheet\.getSheetId\(\)\)/)
  })

  it('every language carries the wording', () => {
    const dir = join(here, '..', 'src/renderer/i18n/app')
    const files = readdirSync(dir).filter((f) => f.endsWith('.ts'))
    expect(files.length).toBeGreaterThanOrEqual(22)
    for (const file of files) {
      expect(readFileSync(join(dir, file), 'utf8'), file).toContain('appShowingFormulasActive')
    }
  })
})

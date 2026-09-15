import { describe, expect, it } from 'vitest'
import {
  capabilityReport,
  familyLabel,
  FORMAT_FAMILIES,
  formatFamily,
  withSaveExtension,
} from '../../src/main/mcp/tools/formats'

/**
 * The format registry is the MCP layer's single source of truth for what each
 * editor can open/save/export vs. what this server exposes. These tests pin the
 * matrix to the shell routing and the per-editor Save-As dialogs so a drift is
 * caught here rather than in a live session.
 */

describe('format registry', () => {
  it('mirrors the shell open routing for the mcp-driven families', () => {
    expect(formatFamily('docx').editorOpen).toEqual(['docx'])
    expect(formatFamily('xlsx').editorOpen).toEqual(['xlsx', 'xlsm', 'xls', 'csv'])
    expect(formatFamily('pptx').editorOpen).toEqual(['pptx'])
  })

  it('keeps the editor-only families (md/html) without an mcp block', () => {
    for (const family of ['md', 'html'] as const) {
      expect(formatFamily(family).mcp).toBeUndefined()
    }
    // only the families MCP actually drives carry an mcp block
    expect(FORMAT_FAMILIES.filter((f) => f.mcp).map((f) => f.family)).toEqual([
      'docx',
      'xlsx',
      'pptx',
    ])
  })

  it('records the editor export capabilities even where mcp cannot reach them', () => {
    expect(formatFamily('docx').editorExport).toEqual(['pdf'])
    expect(formatFamily('md').editorExport).toEqual(['docx', 'pdf'])
    expect(formatFamily('html').editorExport).toEqual(['docx', 'pdf'])
    expect(formatFamily('pdf').editorExport).toEqual(['docx', 'xlsx', 'pptx'])
  })

  it('never advertises an mcp save format the editor cannot write', () => {
    for (const f of FORMAT_FAMILIES) {
      for (const ext of f.mcp?.save ?? []) {
        expect(f.editorSave).toContain(ext)
      }
    }
  })

  it('reports labels for guided errors', () => {
    expect(familyLabel('docx')).toBe('Word document')
    expect(familyLabel('xlsx')).toBe('spreadsheet')
    expect(familyLabel('pptx')).toBe('presentation')
    expect(() => familyLabel('nope' as never)).toThrow(/unknown format family/)
  })
})

describe('withSaveExtension', () => {
  it('keeps a correct family extension, case-insensitively', () => {
    expect(withSaveExtension('docx', 'C:/x/Report.DOCX')).toBe('C:/x/Report.DOCX')
    expect(withSaveExtension('xlsx', '/tmp/books.xlsx')).toBe('/tmp/books.xlsx')
    expect(withSaveExtension('pptx', '/tmp/deck.pptx')).toBe('/tmp/deck.pptx')
  })

  it('appends the family extension when the path has none', () => {
    expect(withSaveExtension('docx', '/tmp/report')).toBe('/tmp/report.docx')
    expect(withSaveExtension('xlsx', '/tmp/books')).toBe('/tmp/books.xlsx')
    expect(withSaveExtension('pptx', '/tmp/deck')).toBe('/tmp/deck.pptx')
  })

  it('rejects a mismatched extension rather than silently rewriting it', () => {
    expect(() => withSaveExtension('xlsx', '/tmp/report.docx')).toThrow(
      /spreadsheet session must be saved as \.xlsx \(got "\.docx"\)/,
    )
    expect(() => withSaveExtension('pptx', '/tmp/deck.docx')).toThrow(/must be saved as \.pptx/)
  })
})

describe('capabilityReport', () => {
  it('exposes editor truth alongside the mcp subset for every family', () => {
    const report = capabilityReport()
    expect(report.map((r) => r.family)).toEqual(['docx', 'xlsx', 'pptx', 'md', 'html', 'pdf'])
    const docx = report.find((r) => r.family === 'docx')!
    expect(docx.editor.open).toEqual(['docx'])
    expect(docx.editor.export).toEqual(['pdf'])
    expect(docx.mcp).toEqual({ generate: 'docx', save: ['docx'], read: 'docx' })
    // families MCP does not drive yet carry no mcp key
    expect(report.find((r) => r.family === 'md')!.mcp).toBeUndefined()
    expect(report.find((r) => r.family === 'pdf')!.mcp).toBeUndefined()
  })
})

import { describe, expect, it } from 'vitest'

import { patchCheckSheetName, type WorkbookProtoLike } from '../src/renderer/sheet-rename-fix'

interface FakeSheet {
  getName(): string
}

/// Mirrors @univerjs/core Workbook: checkSheetName is a case-insensitive
/// duplicate scan reused by the name-minting loops.
class FakeWorkbook {
  private readonly sheets: FakeSheet[]
  private active: FakeSheet | null
  private count = 2

  constructor(names: string[], activeIndex: number | null = 0) {
    this.sheets = names.map((name) => ({ getName: () => name }))
    this.active = activeIndex === null ? null : this.sheets[activeIndex]!
  }

  getActiveSheet(_allowNull: true): FakeSheet | null {
    return this.active
  }

  getSheets(): FakeSheet[] {
    return this.sheets
  }

  checkSheetName(name: string): boolean {
    return this.getSheets().some((sheet) => sheet.getName().toLowerCase() === name.toLowerCase())
  }

  uniqueSheetName(name = 'Sheet1'): string {
    let output = name
    while (this.checkSheetName(output)) {
      output = name + this.count
      this.count += 1
    }
    return output
  }

  generateNewSheetName(name: string): string {
    let output = name + this.count
    while (this.checkSheetName(output)) {
      output = name + this.count
      this.count += 1
    }
    return output
  }
}

function patched(): { dispose(): void } {
  return patchCheckSheetName(FakeWorkbook.prototype as unknown as WorkbookProtoLike)
}

describe('patchCheckSheetName', () => {
  it('allows renaming the active sheet to a case variant of itself', () => {
    const workbook = new FakeWorkbook(['Sheet1', 'Sheet2'])
    const patch = patched()
    try {
      expect(workbook.checkSheetName('SHEET1')).toBe(false)
      expect(workbook.checkSheetName('sheet1')).toBe(false)
    } finally {
      patch.dispose()
    }
  })

  it('still blocks renaming to another existing sheet, any case', () => {
    const workbook = new FakeWorkbook(['Sheet1', 'Sheet2'])
    const patch = patched()
    try {
      expect(workbook.checkSheetName('Sheet2')).toBe(true)
      expect(workbook.checkSheetName('SHEET2')).toBe(true)
    } finally {
      patch.dispose()
    }
  })

  it('keeps the exact active name reported as a duplicate (upstream self-check handles it)', () => {
    const workbook = new FakeWorkbook(['Sheet1'])
    const patch = patched()
    try {
      expect(workbook.checkSheetName('Sheet1')).toBe(true)
    } finally {
      patch.dispose()
    }
  })

  it('does not weaken new-sheet dedupe against a case variant of the active sheet', () => {
    const workbook = new FakeWorkbook(['sheet2', 'Sheet1'], 0)
    const patch = patched()
    try {
      expect(workbook.generateNewSheetName('Sheet')).toBe('Sheet3')
      expect(workbook.uniqueSheetName('SHEET2')).not.toBe('SHEET2')
    } finally {
      patch.dispose()
    }
  })

  it('delegates untouched when there is no active sheet (load path)', () => {
    const workbook = new FakeWorkbook(['Sheet1'], null)
    const patch = patched()
    try {
      expect(workbook.checkSheetName('SHEET1')).toBe(true)
      expect(workbook.checkSheetName('Sheet2')).toBe(false)
    } finally {
      patch.dispose()
    }
  })

  it('restores the original methods on dispose', () => {
    const workbook = new FakeWorkbook(['Sheet1'])
    const original = FakeWorkbook.prototype.checkSheetName
    const patch = patched()
    expect(FakeWorkbook.prototype.checkSheetName).not.toBe(original)
    patch.dispose()
    expect(FakeWorkbook.prototype.checkSheetName).toBe(original)
    expect(workbook.checkSheetName('SHEET1')).toBe(true)
  })
})

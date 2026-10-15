import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

/**
 * `save-keeps-caret-marks.test.ts` covers the helper. These assertions pin the
 * wiring inside saveOnce(), because that is where the defect actually lived:
 * with the restore call removed the whole helper suite still passes, yet the
 * autosave still eats the user's font.
 */

const SRC = join(__dirname, '../src/renderer')
const fileActions = (): string => readFileSync(join(SRC, 'file-actions.ts'), 'utf8')

describe('saveOnce caret-mark wiring', () => {
  it('snapshots the caret marks before the save round-trips the document', () => {
    const src = fileActions()
    const snapshot = src.indexOf('caretMarksSnapshot(editor)')
    // scope to saveOnce: the doc is snapshotted and the bytes built before the
    // off-thread reparse, and the caret marks must be captured with them
    const roundtrip = src.indexOf(
      'await parseDocxOffThread(',
      src.indexOf('async function saveOnce('),
    )
    expect(snapshot).toBeGreaterThan(-1)
    // the snapshot has to exist before the doc can be replaced under it
    expect(snapshot).toBeLessThan(roundtrip)
  })

  it('restores the marks after setContent, inside the rewrite branch only', () => {
    const src = fileActions()
    const rewrite = src.indexOf('if (!unchanged) {')
    const setContent = src.indexOf('editor.commands.setContent(rebasedPm', rewrite)
    const restore = src.indexOf('restoreMarksAfterDocRewrite(editor, caretMarks)', rewrite)
    expect(rewrite).toBeGreaterThan(-1)
    expect(setContent).toBeGreaterThan(rewrite)
    // after the replacement, and before the branch closes — the equal-doc path
    // must stay untouched so undo history, caret and scroll keep surviving
    expect(restore).toBeGreaterThan(setContent)
    const branchEnd = src.indexOf('ctx.setDocCss(', restore)
    expect(restore).toBeLessThan(branchEnd)
  })

  it('keeps annotation marks out of the snapshot', () => {
    // FORMAT_MARKS is the filter; without it a restored comment mark would
    // extend the comment onto whatever the user types next
    const src = fileActions()
    expect(src).toContain('FORMAT_MARKS')
    expect(src).toContain('serializeMarks')
  })
})

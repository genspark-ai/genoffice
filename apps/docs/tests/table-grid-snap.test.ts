/**
 * settings.xml w:compat <w:adjustLineHeightInTable/> + typed w:docGrid
 * (Word probe 2026-09-02, prod-sas 086 package replicas):
 * - with the flag, table-cell lines snap to the grid with full body semantics —
 *   9/10.5/12pt Yu Mincho cells all take one 18pt cell, 14/16/21pt take two;
 *   auto multiples resolve to mult x pitch (276 -> 20.7pt, 360 -> 27pt),
 *   atLeast = max(value, snapped single), exact never snaps, and
 *   w:snapToGrid=0 opts a paragraph out (real title box 23.28pt -> 36pt).
 * - without the flag cells never snap (probe rows stayed at 1.44em), so the
 *   static td/th pitch kill must stay in force.
 * The renderer switches by restoring --doc-grid-pitch inside cells; the
 * existing grid line-height expressions then apply unchanged.
 */
import { describe, expect, it } from 'vitest'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { docStyleCss } from '../src/renderer/doc-style-css'

;(globalThis as { CSS?: unknown }).CSS ??= { escape: (s: string) => s }

const CELL_RULE = '.doc-page .doc-table :is(td, th) { --doc-grid-pitch: inherit }'
const GRID = '<w:docGrid w:type="lines" w:linePitch="360"/>'

const settingsPart = (compatXml: string) => ({
  path: 'word/settings.xml',
  xml:
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:compat>${compatXml}</w:compat></w:settings>`,
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
})

const BODY = '<w:p><w:r><w:t>text</w:t></w:r></w:p>'

describe('adjustLineHeightInTable cell grid snapping', () => {
  it('restores the cell grid pitch when the flag and a typed grid are present', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BODY,
        sectPrExtra: GRID,
        extraParts: [settingsPart('<w:adjustLineHeightInTable/>')],
      }),
    )
    expect(doc.adjustLineHeightInTable).toBe(true)
    expect(docStyleCss(doc)).toContain(CELL_RULE)
  })

  it('keeps cells off the grid without the flag', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: BODY, sectPrExtra: GRID }))
    expect(doc.adjustLineHeightInTable).toBeUndefined()
    expect(docStyleCss(doc)).not.toContain(CELL_RULE)
  })

  it('a val="0" flag counts as off', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BODY,
        sectPrExtra: GRID,
        extraParts: [settingsPart('<w:adjustLineHeightInTable w:val="0"/>')],
      }),
    )
    expect(doc.adjustLineHeightInTable).toBeUndefined()
    expect(docStyleCss(doc)).not.toContain(CELL_RULE)
  })

  it('emits nothing without a typed grid even with the flag', async () => {
    const doc = await parseDocx(
      await buildDocx({
        bodyXml: BODY,
        extraParts: [settingsPart('<w:adjustLineHeightInTable/>')],
      }),
    )
    expect(doc.adjustLineHeightInTable).toBe(true)
    expect(docStyleCss(doc)).not.toContain(CELL_RULE)
  })
})

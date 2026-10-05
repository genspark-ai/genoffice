import assert from 'node:assert/strict'
import JSZip from 'jszip'
import { test } from 'vitest'

import { generateDocx } from '../src/generate'
import { bookmarkName } from '../src/generate/bookmarks'

async function documentXml(ir: unknown[]) {
  const buffer = await generateDocx(ir, {})
  const zip = await JSZip.loadAsync(buffer)
  return (await zip.file('word/document.xml')?.async('string')) ?? ''
}

test('keeps bookmarks on table cells, cards and color bars', async () => {
  const xml = await documentXml([
    {
      type: 'table',
      colWidths: [400],
      rows: [{ cells: [{ runs: [{ text: 'Cell' }], bookmarks: ['cell-anchor'] }] }],
      bookmarks: ['table-anchor'],
    },
    {
      type: 'card',
      children: [{ type: 'para', runs: [{ text: 'Body' }], style: {} }],
      bookmarks: ['card-anchor'],
    },
    {
      type: 'para',
      runs: [{ text: 'Section' }],
      style: { shading: '3355AA', colorBar: true, exactLineHeightPx: 30 },
      bookmarks: ['bar-anchor'],
    },
  ])
  for (const id of ['table-anchor', 'cell-anchor', 'card-anchor', 'bar-anchor']) {
    assert.match(xml, new RegExp(`w:name="${bookmarkName(id)}"`), `missing bookmark ${id}`)
  }
})

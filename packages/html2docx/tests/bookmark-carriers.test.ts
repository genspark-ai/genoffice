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

test('carries a card bookmark into a nested table and through the flattened card paths', async () => {
  const table = {
    type: 'table',
    colWidths: [400],
    rows: [{ cells: [{ runs: [{ text: 'Nested' }] }] }],
  }
  const tall = {
    type: 'card',
    allowSplit: true,
    heightPx: 900,
    shading: 'F5F5F5',
    children: [
      { type: 'para', runs: [{ text: 'A' }], style: {} },
      { type: 'para', runs: [{ text: 'B' }], style: {} },
    ],
    bookmarks: ['flat-anchor'],
  }
  const deep = (children) => ({ type: 'card', children })
  const xml = await documentXml([
    { type: 'card', children: [table], bookmarks: ['nested-table-anchor'] },
    tall,
    deep([
      deep([
        deep([
          {
            type: 'card',
            children: [{ type: 'para', runs: [{ text: 'D' }], style: {} }],
            bookmarks: ['deep-anchor'],
          },
        ]),
      ]),
    ]),
  ])
  for (const id of ['nested-table-anchor', 'flat-anchor', 'deep-anchor']) {
    assert.match(xml, new RegExp(`w:name="${bookmarkName(id)}"`), `missing bookmark ${id}`)
  }
})

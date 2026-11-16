import assert from 'node:assert/strict'
import { test } from 'vitest'

import { normalizeIr, raceWithAbort } from '../src'
import { generateDocx } from '../src/generate'

test('accepts a well-formed IR array', () => {
  const ir = normalizeIr([
    { type: 'paragraph', runs: [] },
    { type: 'image', shotId: 'abc123' },
  ])
  assert.equal(ir.length, 2)
  assert.equal(ir[0].type, 'paragraph')
  assert.equal(ir[1].shotId, 'abc123')
})

test('rejects non-array IR payloads', () => {
  for (const raw of [null, undefined, {}, '[]', 42]) {
    assert.throws(() => normalizeIr(raw), /malformed IR.*expected an array/)
  }
})

test('rejects null and non-object items', () => {
  assert.throws(() => normalizeIr([null]), /index 0/)
  assert.throws(() => normalizeIr(['paragraph']), /index 0/)
  assert.throws(() => normalizeIr([42]), /index 0/)
})

test('rejects items with a missing or empty type', () => {
  assert.throws(() => normalizeIr([{}]), /index 0.*type/)
  assert.throws(() => normalizeIr([{ type: '' }]), /index 0.*type/)
  assert.throws(() => normalizeIr([{ type: 42 }]), /index 0.*type/)
  assert.throws(() => normalizeIr([{ type: 'paragraph' }, null]), /index 1/)
})

test('rejects invalid shotId shapes', () => {
  assert.throws(() => normalizeIr([{ type: 'image', shotId: '' }]), /index 0.*shotId/)
  assert.throws(() => normalizeIr([{ type: 'image', shotId: 42 }]), /index 0.*shotId/)
})

test('rejects non-finite or negative numeric geometry', () => {
  for (const bad of [NaN, Infinity, -5, '-3']) {
    assert.throws(() => normalizeIr([{ type: 'image', width: bad }]), /index 0.*width/)
    assert.throws(() => normalizeIr([{ type: 'image', height: bad }]), /index 0.*height/)
  }
  assert.throws(() => normalizeIr([{ type: 'image', widthFrac: Infinity }]), /widthFrac/)
  assert.throws(() => normalizeIr([{ type: 'image', xPx: NaN }]), /xPx/)
})

test('accepts negative offsets: elements may overhang their origin', () => {
  const ir = normalizeIr([{ type: 'image', xPx: -12, yPx: -3 }])
  assert.equal(ir.length, 1)
})

test('accepts honest numeric geometry', () => {
  const ir = normalizeIr([{ type: 'image', width: 800, height: 600, widthFrac: 0.5 }])
  assert.equal(ir.length, 1)
})

test('rejects a non-finite or negative spacer px, exactly like heightPx', () => {
  // px is the spacer's height: a bad value reaches
  // pxToTwips(Math.max(px, 2)) as w:line="0" with w:lineRule="exact", silently
  // collapsing the intended vertical gap, and poisons __h2dSpacerPx so the
  // >8 and >16 layout fixups in renderer/page-settings stop firing.
  for (const bad of ['not-a-number', NaN, Infinity, -8]) {
    assert.throws(() => normalizeIr([{ type: 'spacer', px: bad }]), /index 0.*px/, `px=${bad}`)
  }
  // identical treatment to the sibling field it is validated alongside
  for (const bad of ['oops', NaN, Infinity, -1]) {
    const onPx = () => normalizeIr([{ type: 'spacer', px: bad }])
    const onSibling = () => normalizeIr([{ type: 'spacer', heightPx: bad }])
    assert.throws(onPx, /index 0/, `px=${bad} should throw`)
    assert.throws(onSibling, /index 0/, `heightPx=${bad} should throw`)
  }
})

test('accepts an honest spacer px and skips absent px the way heightPx does', () => {
  assert.equal(normalizeIr([{ type: 'spacer', px: 24 }])[0].px, 24)
  assert.equal(normalizeIr([{ type: 'spacer', px: 0 }])[0].px, 0)
  // undefined/null are absent, not invalid — unchanged from the siblings
  assert.equal(normalizeIr([{ type: 'spacer' }]).length, 1)
  assert.equal(normalizeIr([{ type: 'spacer', px: null }]).length, 1)
})

test('a rejected spacer px never reaches the generated geometry', async () => {
  // Before the fix this produced w:line="0" (an exact zero-height paragraph);
  // now the hostile value is refused at the same boundary as every other key.
  assert.throws(
    () =>
      normalizeIr([
        { type: 'spacer', px: 'not-a-number' },
        { type: 'para', runs: [] },
      ]),
    /invalid numeric "px"/,
  )
})

test('rejects a paragraph, heading, list or table the generator cannot walk', () => {
  // Page JavaScript is the source of this IR (see normalizeIr's contract), and
  // these fields are dereferenced without a fallback: `node.runs.length`,
  // `item.runs[0]`, `row.cells.map`, and HEADING_LEVELS[node.level - 1].
  const bad: Array<[unknown, RegExp]> = [
    [{ type: 'para' }, /index 0.*missing array "runs"/],
    [{ type: 'para', runs: null }, /missing array "runs"/],
    [{ type: 'heading', runs: [] }, /missing numeric "level"/],
    [{ type: 'heading', level: null, runs: [] }, /missing numeric "level"/],
    [{ type: 'heading', level: 'x', runs: [] }, /missing numeric "level"/],
    [{ type: 'para', level: 2 }, /missing array "runs"/],
    [{ type: 'list' }, /missing array "items"/],
    [{ type: 'list', items: [{ runs: [] }, {}] }, /list item 1 missing array "runs"/],
    [{ type: 'list', items: [{ style: {} }] }, /list item 0 missing array "runs"/],
    [{ type: 'table' }, /missing array "rows"/],
    [{ type: 'table', rows: [{}] }, /table row 0 missing array "cells"/],
    [
      { type: 'table', rows: [{ cells: [] }, { cells: null }] },
      /table row 1 missing array "cells"/,
    ],
  ]
  for (const [node, message] of bad) {
    assert.throws(() => normalizeIr([node]), message, JSON.stringify(node))
  }
})

test('keeps accepting the shapes the extractor writes', () => {
  const ir = normalizeIr([
    { type: 'para', runs: [] },
    { type: 'heading', level: 2, runs: [{ text: 'Title' }] },
    { type: 'list', items: [[{ text: 'old snapshot shape' }], { runs: [{ text: 'current' }] }] },
    { type: 'table', rows: [{ cells: [] }, { cells: [{ runs: [{ text: 'cell' }] }] }] },
    { type: 'image', shotId: 'abc123' },
  ])
  assert.equal(ir.length, 5)
})

test('rejects the shapes a compromised page can inject, which the generator cannot walk', async () => {
  // The extractor calls into `globalThis.__html2docxTable` and friends, which
  // classic scripts install on the page — page JavaScript can replace them, so
  // these shapes are reachable IR even though the shipped extractor never
  // writes them. Both throw a bare TypeError from inside the renderer
  // (`node.runs.length`, `row.cells.map`) unless the boundary refuses them.
  await assert.rejects(generateDocx([{ type: 'para' }] as never, {}), TypeError)
  await assert.rejects(
    generateDocx([{ type: 'table', rows: [{ cells: null }] }] as never, {}),
    TypeError,
  )
  assert.throws(() => normalizeIr([{ type: 'para' }]), /index 0.*missing array "runs"/)
  assert.throws(
    () => normalizeIr([{ type: 'table', rows: [{ cells: null }] }]),
    /table row 0 missing array "cells"/,
  )
})

test('raceWithAbort resolves when there is no signal', async () => {
  const value = await raceWithAbort(Promise.resolve('ok'))
  assert.equal(value, 'ok')
})

test('raceWithAbort rejects immediately when already aborted', async () => {
  const controller = new AbortController()
  controller.abort()
  await assert.rejects(raceWithAbort(new Promise(() => {}), controller.signal), /aborted/)
})

test('raceWithAbort rejects when abort fires during a pending wait', async () => {
  const controller = new AbortController()
  const pending = new Promise<string>((resolve) => setTimeout(() => resolve('late'), 5000))
  const raced = raceWithAbort(pending, controller.signal)
  controller.abort()
  await assert.rejects(raced, /aborted/)
})

test('raceWithAbort propagates the original rejection', async () => {
  const controller = new AbortController()
  await assert.rejects(raceWithAbort(Promise.reject(new Error('boom')), controller.signal), /boom/)
})

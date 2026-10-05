import assert from 'node:assert/strict'
import { test } from 'vitest'

import { normalizeIr, raceWithAbort } from '../src'

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

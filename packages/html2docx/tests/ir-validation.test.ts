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

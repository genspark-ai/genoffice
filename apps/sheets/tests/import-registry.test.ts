import { describe, expect, it } from 'vitest'

import { ImportRegistry } from '../src/main/import-registry'

/**
 * The cancel for a delimited import (a .csv/.tsv opened as a workbook) has to
 * travel from the renderer to the main process and reach the work it is
 * meant to stop. This registry is where the two meet, so it is where the
 * wiring is checked: a cancel that reaches nothing leaves the user waiting
 * on a 250 MB file with a button that does nothing.
 */
describe('the import registry', () => {
  it('gives the running import a controller the caller can watch', () => {
    const registry = new ImportRegistry()
    const controller = registry.start(7)
    expect(controller.signal.aborted).toBe(false)
    expect(registry.controllerFor(7)).toBe(controller)
  })

  it('cancels the import running for that owner', () => {
    const registry = new ImportRegistry()
    const controller = registry.start(7)
    expect(registry.cancel(7)).toBe(true)
    expect(controller.signal.aborted).toBe(true)
  })

  it('says nothing happened when there is no import to cancel', () => {
    const registry = new ImportRegistry()
    expect(registry.cancel(7)).toBe(false)
    expect(registry.controllerFor(7)).toBeUndefined()
  })

  it('cancels only the owner that asked', () => {
    const registry = new ImportRegistry()
    const mine = registry.start(1)
    const theirs = registry.start(2)
    registry.cancel(1)
    expect(mine.signal.aborted).toBe(true)
    expect(theirs.signal.aborted).toBe(false)
  })

  it('does not cancel a finished import', () => {
    const registry = new ImportRegistry()
    const controller = registry.start(7)
    registry.finish(7, controller)
    expect(registry.controllerFor(7)).toBeUndefined()
    expect(registry.cancel(7)).toBe(false)
    expect(controller.signal.aborted).toBe(false)
  })

  it('lets a second import replace the first', () => {
    const registry = new ImportRegistry()
    const first = registry.start(7)
    const second = registry.start(7)
    expect(registry.controllerFor(7)).toBe(second)
    // the first import's work is stale, but it must not be cancelled by a
    // cancel meant for the second one
    registry.cancel(7)
    expect(second.signal.aborted).toBe(true)
    expect(first.signal.aborted).toBe(false)
  })

  it('does not let a slow import clear the entry of the next one', () => {
    const registry = new ImportRegistry()
    const first = registry.start(7)
    const second = registry.start(7)
    // the first import finishes late, after the second one started
    expect(registry.finish(7, first)).toBe(false)
    expect(registry.controllerFor(7)).toBe(second)
    expect(registry.cancel(7)).toBe(true)
    expect(second.signal.aborted).toBe(true)
    expect(registry.finish(7, second)).toBe(true)
    expect(registry.controllerFor(7)).toBeUndefined()
  })

  it('tells a cancelled import from one that failed on its own', () => {
    const registry = new ImportRegistry()
    const controller = registry.start(7)
    expect(registry.wasCancelled(controller)).toBe(false)
    registry.cancel(7)
    expect(registry.wasCancelled(controller)).toBe(true)
  })

  it('does not call a failed import cancelled', () => {
    // An import that failed on its own (a file too large, a broken archive)
    // has to report the reason; only the user's own cancel opens nothing and
    // says nothing.
    const registry = new ImportRegistry()
    const controller = registry.start(7)
    registry.finish(7, controller)
    expect(registry.wasCancelled(controller)).toBe(false)
  })
})

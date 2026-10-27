/**
 * Sort & Filter ▸ Filter must reach a command that exists.
 *
 * Univer defines `sheet.command.smart-toggle-filter` and points its own menu
 * and its ⇧⌘L shortcut at it, but `UniverSheetsFilterPlugin` never registers
 * it — the plugin registers the five commands around it and skips the toggle.
 * A click therefore asked the command service for a command that was not
 * there: nothing ran, nothing was said, and the sheet kept exactly the state
 * it had, which reads as a dead button.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import type { ICommandService } from '@univerjs/core'
import { SmartToggleSheetsFilterCommand } from '@univerjs/sheets-filter'
import { registerFilterToggleCommand } from '../src/renderer/filter-toggle-command-fix'

const here = dirname(fileURLToPath(import.meta.url))
const read = (rel: string) => readFileSync(join(here, '..', rel), 'utf8')

const TOGGLE_ID = 'sheet.command.smart-toggle-filter'

/** A command service that knows only what it was told. */
class FakeCommandService {
  readonly commands = new Map<string, unknown>()
  hasCommand(id: string): boolean {
    return this.commands.has(id)
  }
  registerCommand(command: { id: string }): { dispose(): void } {
    this.commands.set(command.id, command)
    return {
      dispose: () => {
        this.commands.delete(command.id)
      },
    }
  }
}

const asService = (service: FakeCommandService): ICommandService =>
  service as unknown as ICommandService

describe('the filter-toggle command', () => {
  it('is the command the ribbon asks for', () => {
    expect(SmartToggleSheetsFilterCommand.id).toBe(TOGGLE_ID)
    // upstream's handler is what does the work; installing the command is
    // only ever a matter of making it reachable
    expect(typeof SmartToggleSheetsFilterCommand.handler).toBe('function')
  })

  it('is registered, so a click reaches a handler instead of nothing', () => {
    const service = new FakeCommandService()
    registerFilterToggleCommand(asService(service))
    expect(service.hasCommand(TOGGLE_ID)).toBe(true)
  })

  it('leaves a command that is already registered alone', () => {
    const service = new FakeCommandService()
    const own = { id: TOGGLE_ID, handler: () => true }
    service.commands.set(own.id, own)
    registerFilterToggleCommand(asService(service))
    expect(service.commands.get(own.id)).toBe(own)
  })

  it('unregisters when the editor goes away', () => {
    const service = new FakeCommandService()
    const disposable = registerFilterToggleCommand(asService(service))
    disposable.dispose()
    expect(service.hasCommand(TOGGLE_ID)).toBe(false)
  })

  it('is installed when the editor starts', () => {
    expect(read('src/renderer/App.tsx')).toContain('installFilterToggleCommand(runtime)')
  })

  it('says so when the command refuses, instead of leaving the sheet unchanged in silence', () => {
    const src = read('src/renderer/ribbon-actions.ts')
    expect(src).toContain(`executeCommand('${TOGGLE_ID}')`)
    expect(src).toContain('appFilterToggleFailed')
  })
})

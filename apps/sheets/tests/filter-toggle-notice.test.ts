/**
 * A declined Filter toggle must not be silent.
 *
 * `sheet.command.smart-toggle-filter` returns false when it refuses — no
 * active selection, a range it cannot take — and the ribbon used to drop that
 * answer on the floor. The sheet kept exactly the state it had, no filter
 * arrows appeared, and nothing said why: the button read as dead (user
 * report, genoffice#2243).
 *
 * While the sheet is still loading, the gate in App.tsx has its own answer
 * (a footer note, a "load it all" prompt, or the column-projection path's
 * progress), so the ribbon stays quiet rather than talking over it.
 */
import { describe, expect, it } from 'vitest'
import { handleRibbonCommand, type RibbonCommandContext } from '../src/renderer/ribbon-actions'

interface Harness {
  ctx: RibbonCommandContext
  messages: string[]
}

function makeFilterHarness(options: { ok: boolean; preloadComplete?: boolean }): Harness {
  const messages: string[] = []
  const ctx = {
    univerRef: {
      current: {
        univerAPI: {
          executeCommand: () => Promise.resolve(options.ok),
          getActiveWorkbook: () => undefined,
          getActiveSheet: () => undefined,
        },
      },
    },
    lazyWorkbookRef: {
      current: { flags: { preloadComplete: options.preloadComplete ?? true } },
    },
    setMessage: (message: string) => {
      messages.push(message)
    },
  } as unknown as RibbonCommandContext
  return { ctx, messages }
}

/** let the ribbon's own `then` run */
const settle = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0))

describe('the Filter toggle', () => {
  it('says so when the command declines', async () => {
    const { ctx, messages } = makeFilterHarness({ ok: false })
    handleRibbonCommand(ctx, 'filter-toggle')
    await settle()
    expect(messages).toHaveLength(1)
    expect(messages[0]).toMatch(/filter|筛选/i)
  })

  it('stays quiet when the filter went on', async () => {
    const { ctx, messages } = makeFilterHarness({ ok: true })
    handleRibbonCommand(ctx, 'filter-toggle')
    await settle()
    expect(messages).toHaveLength(0)
  })

  it('stays quiet while the sheet is still loading, because the gate says why', async () => {
    const { ctx, messages } = makeFilterHarness({ ok: false, preloadComplete: false })
    handleRibbonCommand(ctx, 'filter-toggle')
    await settle()
    expect(messages).toHaveLength(0)
  })
})

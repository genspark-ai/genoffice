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
 * progress), so the ribbon stays quiet rather than talking over it. A
 * canceling protection guard also answers false, and it has already said
 * why — the toggle stays quiet there too rather than overwriting it.
 */
import { describe, expect, it } from 'vitest'
import { handleRibbonCommand, type RibbonCommandContext } from '../src/renderer/ribbon-actions'
import {
  UNPROTECTED_SHEET,
  mergeSheetProtectionAllow,
  type SheetProtectionInfo,
} from '../src/renderer/sheet-protection'

interface Harness {
  ctx: RibbonCommandContext
  messages: string[]
}

function makeFilterHarness(options: {
  ok: boolean
  preloadComplete?: boolean
  /** the effective protection of the active sheet; `'unknown'` = not indexed yet */
  sheetProtection?: SheetProtectionInfo | 'unknown'
}): Harness {
  const messages: string[] = []
  const sheetId = 'sh1'
  const protection =
    options.sheetProtection === undefined ? UNPROTECTED_SHEET : options.sheetProtection
  const state = {
    flags: { preloadComplete: options.preloadComplete ?? true },
    sheetProtections:
      protection && protection !== 'unknown' ? new Map([[sheetId, protection]]) : new Map(),
    editJournal: { sheetProtection: new Map(), sheets: { added: new Set() } },
  }
  const ctx = {
    univerRef: {
      current: {
        univerAPI: {
          executeCommand: () => Promise.resolve(options.ok),
          getActiveWorkbook: () => ({ getActiveSheet: () => ({ getSheetId: () => sheetId }) }),
          getActiveSheet: () => undefined,
        },
      },
    },
    lazyWorkbookRef: { current: state },
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
    expect(messages[0]).toMatch(/filter|\u7b5b\u9009/i)
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

  it('stays quiet when the protection guard blocked the command, because the guard says why', async () => {
    // The guard cancels through the same `false` and has already toasted
    // `appProtectedSheetCommandBlocked`; the toggle's own advice (select a
    // range with data first) would overwrite it with the wrong reason.
    const { ctx, messages } = makeFilterHarness({
      ok: false,
      sheetProtection: {
        protected: true,
        hasPassword: false,
        allow: mergeSheetProtectionAllow({ autoFilter: false }),
        password: null,
      },
    })
    handleRibbonCommand(ctx, 'filter-toggle')
    await settle()
    expect(messages).toHaveLength(0)
  })

  it('stays quiet while the protection is still unknown, because the guard says why', async () => {
    // Not indexed yet: the guard fails closed with `appProtectionNeedsIndexed`.
    const { ctx, messages } = makeFilterHarness({ ok: false, sheetProtection: 'unknown' })
    handleRibbonCommand(ctx, 'filter-toggle')
    await settle()
    expect(messages).toHaveLength(0)
  })

  it('speaks on a genuine decline of a protected sheet that allows filters', async () => {
    const { ctx, messages } = makeFilterHarness({
      ok: false,
      sheetProtection: {
        protected: true,
        hasPassword: false,
        allow: mergeSheetProtectionAllow({ autoFilter: true }),
        password: null,
      },
    })
    handleRibbonCommand(ctx, 'filter-toggle')
    await settle()
    expect(messages).toHaveLength(1)
  })
})

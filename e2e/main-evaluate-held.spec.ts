/**
 * `app.evaluate` in the main process survives a collection that runs before
 * the main process gets to its microtasks. The order is forced: the call is
 * made to land inside JavaScript that is still running, which collects before
 * it returns. Unwrapped, such a call fails with "Resulting promise was
 * garbage collected".
 */
import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo } from './helpers'

test('a main-process evaluation that lands inside running JavaScript is not lost to a collection', async () => {
  const launched = await launchShell({ onboardingSeen: true, videoDir: 'main-evaluate-held' })
  try {
    const { app } = launched
    await app.evaluate(() => {
      const scope = globalThis as { __evaluated?: boolean; __collectedAfter?: number }
      scope.__collectedAfter = 0
      process.getBuiltinModule('v8').setFlagsFromString('--expose-gc')
      const collect = process.getBuiltinModule('vm').runInNewContext('gc') as () => void
      setInterval(() => {
        const end = Date.now() + 45
        while (Date.now() < end) {
          if (!scope.__evaluated) continue
          scope.__evaluated = false
          scope.__collectedAfter!++
          collect()
        }
        scope.__evaluated = false
      }, 50)
    })
    const failures: string[] = []
    for (let i = 0; i < 150; i++) {
      await app
        .evaluate(() => {
          ;(globalThis as { __evaluated?: boolean }).__evaluated = true
        })
        .catch((error) => failures.push(String(error)))
    }
    expect(failures).toEqual([])
    // the forced order did happen
    const forced = await app.evaluate(
      () => (globalThis as { __collectedAfter?: number }).__collectedAfter,
    )
    expect(forced).toBeGreaterThan(0)
    expect(
      await app.evaluate(async ({ app: electronApp }, n) => electronApp.getName().length + n, 0),
    ).toBeGreaterThan(0)
    await expect(app.evaluate(() => Promise.reject(new Error('refused')))).rejects.toThrow(
      'refused',
    )
    await expect(
      app.evaluate(() => {
        throw new Error('thrown')
      }),
    ).rejects.toThrow('thrown')
  } finally {
    await closeAndSaveVideo(launched, 'main-evaluate-held')
  }
})

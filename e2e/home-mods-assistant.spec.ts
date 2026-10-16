import { test } from '@playwright/test'
import { launchShell } from './helpers'

async function probe(label: string, work: Promise<unknown>, ms = 4000): Promise<string> {
  const result = await Promise.race([
    work.then(() => 'ok').catch((e) => 'err:' + String((e as Error).message).slice(0, 60)),
    new Promise<string>((r) => setTimeout(() => r('TIMEOUT'), ms)),
  ])
  console.log(label, result)
  return result
}

test.describe('probe', () => {
  test('is the app responsive before and after the panel mounts', async () => {
    test.setTimeout(240_000)
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'probe' })
    const { page, app } = launched
    for (let i = 0; i < 8; i++) {
      const r = await probe(
        'boot-' + i,
        page.evaluate(() => 1),
      )
      if (r !== 'ok') break
      await page.waitForTimeout(2000)
    }
    console.log('--- clicking the panel ---')
    await probe(
      'click',
      page.evaluate(() => {
        const el = [...document.querySelectorAll('.nav-item')].find((n) =>
          (n.textContent ?? '').includes('Feature assistant'),
        )
        ;(el as HTMLElement | undefined)?.click()
        return true
      }),
    )
    for (let i = 0; i < 6; i++) {
      await probe(
        'after-' + i,
        page.evaluate(() => document.querySelectorAll('.mods-assistant').length),
      )
      await probe(
        'after-main-' + i,
        app.evaluate(() => 1),
      )
      await page.waitForTimeout(2000)
    }
  })
})

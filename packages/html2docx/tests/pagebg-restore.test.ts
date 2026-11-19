import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import { afterAll, beforeAll, expect, test } from 'vitest'
import type { Browser } from 'playwright-core'

import { convertHtmlToDocx } from '../src'
import { launchChrome, PlaywrightDriver } from '../src/drivers/playwright'

const A4 = { width: 794, height: 1123, deviceScaleFactor: 2 }
let browser: Browser

beforeAll(async () => {
  browser = await launchChrome()
})
afterAll(async () => {
  await browser?.close()
})

// A gradient body background is what makes the extractor build a page
// backdrop screenshot, which is the only path that stamps body children with
// data-h2d-old-visibility. The script appends a hidden body child in the
// window between the set phase and the restore phase: the observer fires only
// on the set phase's inline z-index 2147483647, and MutationObserver callbacks
// are microtasks, so the append lands after the set loop finishes and before
// the restore loop runs.
const PAGE = `<!doctype html><html><head><style>
  body { background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); padding: 24px; }
  .content-box { background: white; border-radius: 8px; padding: 20px; margin: 20px 0;
    box-shadow: 0 2px 10px rgba(0,0,0,.1); border: 1px solid rgba(0,0,0,.08); }
</style></head><body>
  <div class="content-box">The first party agrees to provide comprehensive maintenance
  services covering documentation, training, and on-site inspection within two days.</div>
  <div class="content-box">The second party shall maintain strict confidentiality regarding
  all proprietary business information under this cooperation agreement letter.</div>
  <script>
    window.__probe = null
    const obs = new MutationObserver(() => {
      const stamped = [...document.body.children].some(
        (c) => c.style.zIndex === '2147483647' && c.style.position === 'fixed',
      )
      if (!stamped || window.__probe) return
      const probe = document.createElement('div')
      probe.id = 'late-probe'
      // the page's own intent: this overlay is hidden and must stay hidden
      probe.style.visibility = 'hidden'
      probe.textContent = 'late node'
      document.body.appendChild(probe)
      window.__probe = 'injected'
    })
    obs.observe(document.body, { attributes: true, childList: true, subtree: true })
  </script>
</body></html>`

test('pagebg restore leaves a body child added mid-window untouched', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'html2docx-pagebg-'))
  const input = path.join(dir, 'page.html')
  await fs.writeFile(input, PAGE)
  const driver = await PlaywrightDriver.create(browser, A4)
  try {
    const result = await convertHtmlToDocx({ url: pathToFileURL(input).href }, driver)
    // Guard against a vacuous pass: the pagebg shot path must have run, and
    // the probe must actually have landed inside the set/restore window.
    const pagebg = (result.ir as any[]).find((node) => node.type === 'pagebg')
    assert.ok(pagebg?.shotId, 'page produced a page background screenshot')

    const observed = await driver.evaluate<{
      injected: string | null
      probeVisibility: string | null
      probeAttribute: string | null
      leftoverStamps: number
    }>(() => {
      const probe = document.getElementById('late-probe') as HTMLElement | null
      return {
        injected: (window as any).__probe ?? null,
        probeVisibility: probe ? probe.style.visibility : null,
        probeAttribute: probe ? probe.getAttribute('data-h2d-old-visibility') : null,
        leftoverStamps: document.querySelectorAll('[data-h2d-old-visibility]').length,
      }
    })

    expect(observed.injected).toBe('injected')
    // The node was never stamped, so the restore must not have written to it.
    // Before the fix this read as '' — the page's own 'hidden' was clobbered.
    expect(observed.probeAttribute).toBeNull()
    expect(observed.probeVisibility).toBe('hidden')
    // Stamped nodes are still cleaned up.
    expect(observed.leftoverStamps).toBe(0)
  } finally {
    await driver.close()
  }
})

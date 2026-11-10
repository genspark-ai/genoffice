import { test, expect } from '@playwright/test'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { closeAndSaveVideo, launchShell, screenshotPath } from './helpers'

/**
 * The custom-endpoint profile pane: a self-hosted gateway is described in the
 * settings UI (protocol, headers, capability switches, timeouts, credential
 * indirection, per-model declarations), the description survives a save, and
 * it is what the request path reads back.
 */

const BASE_URL = 'https://gw.example.com/v1'

/** Open Settings → AI Model and select the Custom provider. */
async function openCustomProvider(page: import('@playwright/test').Page) {
  // the star prompt re-appears on later launches and would swallow the first click
  const starred = page.getByRole('button', { name: 'Already starred', exact: true })
  if (await starred.isVisible().catch(() => false)) await starred.click()
  await page.getByRole('button', { name: 'Settings', exact: true }).click()
  await page.locator('.set-nav-item').filter({ hasText: 'AI Model' }).click()
  await page.getByRole('button', { name: 'Provider', exact: true }).click()
  await page.getByRole('option', { name: 'Custom', exact: true }).click()
}

test('a custom endpoint carries a full declarative profile through save and relaunch', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { starPrompt: { resolved: true } },
    videoDir: 'settings-endpoint-profile',
  })
  const { page, userDataDir } = launched
  try {
    await openCustomProvider(page)

    // a fresh settings file has no endpoint yet: the profile pane appears with one
    await page.getByRole('button', { name: 'Add endpoint', exact: true }).click()
    await page.locator('#set-ai-base-url').fill(BASE_URL)
    await page.locator('#set-ai-model').fill('step-3.7-flash')

    // ── the profile pane is present and every control is reachable ──
    const protocol = page.getByRole('button', { name: 'API protocol', exact: true })
    await expect(protocol).toBeVisible()
    await expect(page.getByText('Endpoint capabilities')).toBeVisible()

    // ── declare the Anthropic wire protocol ──
    await protocol.click()
    await page.getByRole('option', { name: 'Anthropic Messages', exact: true }).click()
    await expect(protocol).toHaveText('Anthropic Messages')

    // ── the tri-state switches carry a declaration ──
    for (const name of [
      'Omit the temperature field',
      'Echo reasoning_content',
      'Use max_completion_tokens',
    ]) {
      await page
        .getByRole('radiogroup', { name, exact: true })
        .getByRole('radio', { name: 'On', exact: true })
        .click()
    }

    // ── image input is a tri-state dropdown; "Not supported" declares text-only ──
    await page.getByRole('button', { name: 'Image input', exact: true }).click()
    await page.getByRole('option', { name: 'Not supported', exact: true }).click()

    // ── headers, idle timeout, key env, model declarations ──
    await page.locator('#set-ai-endpoint-headers').fill('X-Tenant: office\nX-Route: eu')
    await page.locator('#set-ai-endpoint-idle').fill('300')
    await page.locator('#set-ai-endpoint-key-env').fill('GW_API_KEY')
    await page
      .locator('#set-ai-endpoint-models')
      .fill('[\n  {\n    "id": "step-3.7-flash",\n    "maxTokens": 8192\n  }\n]')
    // blur the blob so it commits
    await page.locator('#set-ai-endpoint-idle').focus()

    await page.screenshot({ path: screenshotPath('settings-endpoint-profile-pane') })

    await page.getByRole('button', { name: 'Save', exact: true }).click()
    await expect(page.locator('.set-pane-actions .set-ai-status.ok')).toBeVisible()

    // ── the saved file carries the whole profile ──
    const saved = JSON.parse(await readFile(join(userDataDir, 'ai-settings.json'), 'utf8'))
    const ep = saved.customEndpoints[0]
    expect(saved.provider).toBe('custom')
    expect(ep.baseUrl).toBe(BASE_URL)
    expect(ep.model).toBe('step-3.7-flash')
    expect(ep.api).toBe('anthropic')
    expect(ep.headers).toEqual({ 'X-Tenant': 'office', 'X-Route': 'eu' })
    expect(ep.idleTimeoutMs).toBe(300000)
    expect(ep.apiKeyEnv).toBe('GW_API_KEY')
    expect(ep.compat).toEqual({
      omitTemperature: true,
      echoReasoning: true,
      useMaxCompletionTokens: true,
      vision: false,
    })
    expect(ep.modelProfiles).toEqual([{ id: 'step-3.7-flash', maxTokens: 8192 }])

    // ── and the mirrored slot the request path reads ──
    expect(saved.providers.custom.baseUrl).toBe(BASE_URL)
    expect(saved.providers.custom.profile.api).toBe('anthropic')
  } finally {
    await closeAndSaveVideo(launched, 'settings-endpoint-profile')
  }

  // ── relaunch: the pane reflects the stored profile, not a blank form ──
  // The first instance is closed before the second starts: both share one
  // userDataDir, and the single-instance lock makes a concurrent second
  // launch exit immediately (its window closes before firstWindow resolves).
  const again = await launchShell({
    userDataDir,
    onboardingSeen: true,
    videoDir: 'settings-endpoint-profile-reload',
  })
  try {
    await openCustomProvider(again.page)
    await expect(again.page.getByRole('button', { name: 'API protocol', exact: true })).toHaveText(
      'Anthropic Messages',
    )
    await expect(again.page.locator('#set-ai-endpoint-idle')).toHaveValue('300')
    await expect(again.page.locator('#set-ai-endpoint-key-env')).toHaveValue('GW_API_KEY')
    await expect(again.page.locator('#set-ai-endpoint-models')).toHaveValue(/maxTokens/)
    await again.page.screenshot({ path: screenshotPath('settings-endpoint-profile-reload') })
  } finally {
    await closeAndSaveVideo(again, 'settings-endpoint-profile-reload')
  }
})

test('a malformed header line is refused and the reason is shown', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { starPrompt: { resolved: true } },
    videoDir: 'settings-endpoint-profile-bad',
  })
  const { page } = launched
  try {
    await openCustomProvider(page)
    await page.getByRole('button', { name: 'Add endpoint', exact: true }).click()
    await page.locator('#set-ai-base-url').fill(BASE_URL)
    await page.locator('#set-ai-model').fill('m')

    await page.locator('#set-ai-endpoint-headers').fill('X-Tenant office')
    await page.locator('#set-ai-endpoint-idle').focus()
    await expect(
      page.getByText('Invalid header format: each line must be "Name: Value".'),
    ).toBeVisible()
    await expect(page.locator('.set-input-invalid')).toHaveCount(1)

    // fixing the line clears the error
    await page.locator('#set-ai-endpoint-headers').fill('X-Tenant: office')
    await page.locator('#set-ai-endpoint-idle').focus()
    await expect(
      page.getByText('Invalid header format: each line must be "Name: Value".'),
    ).toHaveCount(0)
  } finally {
    await launched.app.close()
  }
})

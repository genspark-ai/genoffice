import { test, expect } from '@playwright/test'
import { createServer, type Server } from 'node:http'
import type { AddressInfo } from 'node:net'
import { closeAndSaveVideo, launchShell, screenshotPath } from './helpers'

/**
 * The self-hosted-gateway loop, against a local server that answers in the
 * shape a declared protocol names: the settings pane probes it on that
 * protocol's route, the picker fills from what it answers, and the connection
 * test reaches it with the right auth header and body. Two protocols are
 * covered — the historic OpenAI-compatible default (a pre-profile endpoint
 * must behave byte for byte as before) and the Anthropic Messages shape a
 * domestic gateway publishes.
 */

interface Seen {
  probePath?: string
  probeAuth?: string
  chatPath?: string
  chatAuth?: string
  chatBody?: Record<string, unknown>
}

let server: Server
let base: string
let seen: Seen

test.beforeAll(async () => {
  server = createServer((req, res) => {
    const chunks: Buffer[] = []
    req.on('data', (c: Buffer) => chunks.push(c))
    req.on('end', () => {
      const send = (value: unknown) => {
        res.writeHead(200, { 'content-type': 'application/json' })
        res.end(JSON.stringify(value))
      }
      const auth = req.headers.authorization ?? req.headers['x-api-key']
      if (req.method === 'GET' && req.url?.endsWith('/models')) {
        seen.probePath = req.url
        seen.probeAuth = Array.isArray(auth) ? auth[0] : auth
        send({ data: [{ id: 'local-model-a' }, { id: 'local-model-b' }], has_more: false })
        return
      }
      if (req.method === 'POST') {
        seen.chatPath = req.url
        seen.chatAuth = Array.isArray(auth) ? auth[0] : auth
        seen.chatBody = JSON.parse(Buffer.concat(chunks).toString('utf8'))
        if (req.url?.endsWith('/messages')) {
          send({
            id: 'msg_local',
            type: 'message',
            role: 'assistant',
            content: [{ type: 'text', text: 'OK' }],
            stop_reason: 'end_turn',
          })
          return
        }
        send({ choices: [{ message: { content: 'OK' }, finish_reason: 'stop' }] })
        return
      }
      res.writeHead(404, { 'content-type': 'application/json' })
      res.end(JSON.stringify({ error: 'not found' }))
    })
  })
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
  base = `http://127.0.0.1:${(server.address() as AddressInfo).port}/v1`
})

test.afterAll(async () => {
  await new Promise<void>((resolve) => server.close(() => resolve()))
})

test.beforeEach(() => {
  seen = {}
})

/** Drive the pane: custom provider → endpoint → base URL → key → protocol → model → test. */
async function drive(
  page: import('@playwright/test').Page,
  protocol: 'openai-compatible' | 'anthropic',
) {
  await page.getByRole('button', { name: 'Settings', exact: true }).click()
  await page.locator('.set-nav-item').filter({ hasText: 'AI Model' }).click()
  await page.getByRole('button', { name: 'Provider', exact: true }).click()
  await page.getByRole('option', { name: 'Custom', exact: true }).click()
  await page.getByRole('button', { name: 'Add endpoint', exact: true }).click()
  await page.locator('#set-ai-base-url').fill(base)
  await page.getByLabel('API Key').fill('local-key')
  if (protocol !== 'openai-compatible') {
    await page.getByRole('button', { name: 'API protocol', exact: true }).click()
    await page.getByRole('option', { name: 'Anthropic Messages', exact: true }).click()
  }
  const modelPicker = page.getByRole('button', { name: 'Model', exact: true })
  await expect(modelPicker).toBeVisible({ timeout: 10_000 })
  await modelPicker.click()
  await page.getByRole('option', { name: 'local-model-b', exact: true }).click()
  await page.getByRole('button', { name: 'Test connection', exact: true }).click()
  await expect(page.locator('.set-pane-actions .set-ai-status.ok')).toHaveText('Connection OK', {
    timeout: 20_000,
  })
}

test('the default OpenAI-compatible shape still probes and connects as it always did', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { starPrompt: { resolved: true } },
    videoDir: 'settings-endpoint-openai-compat',
  })
  const { page } = launched
  try {
    await drive(page, 'openai-compatible')
    // the historic routes and auth header, unchanged
    expect(seen.probePath).toBe('/v1/models')
    expect(seen.probeAuth).toBe('Bearer local-key')
    expect(seen.chatPath).toBe('/v1/chat/completions')
    expect(seen.chatAuth).toBe('Bearer local-key')
    // the one-shot OpenAI path carries model + temperature and no cap field,
    // exactly the body chatOpenAiCompatible has always sent
    expect(seen.chatBody).toMatchObject({ model: 'local-model-b', temperature: 0.3 })
    expect(seen.chatBody?.max_tokens).toBeUndefined()
    await page.screenshot({ path: screenshotPath('settings-endpoint-openai-compat') })
  } finally {
    await closeAndSaveVideo(launched, 'settings-endpoint-gateway')
  }
})

test('a declared Anthropic gateway is probed and connected on its own route', async () => {
  const launched = await launchShell({
    onboardingSeen: true,
    settings: { starPrompt: { resolved: true } },
    videoDir: 'settings-endpoint-local-gateway',
  })
  const { page } = launched
  try {
    await drive(page, 'anthropic')
    // the probe rides the same composed base the request uses (no /v1 doubling)
    expect(seen.probePath).toBe('/v1/models')
    expect(seen.probeAuth).toBe('local-key')
    expect(seen.chatPath).toBe('/v1/messages')
    expect(seen.chatAuth).toBe('local-key')
    expect(seen.chatBody).toMatchObject({ model: 'local-model-b', max_tokens: 8192 })
    await page.screenshot({ path: screenshotPath('settings-endpoint-local-gateway') })
  } finally {
    await closeAndSaveVideo(launched, 'settings-endpoint-gateway')
  }
})

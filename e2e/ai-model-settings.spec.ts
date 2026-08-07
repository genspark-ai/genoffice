/**
 * Drives the AI model settings dialog against a stub OpenAI-compatible server,
 * then checks that what the dialog saved is what every editor module would
 * actually call — the file on disk and the resolved settings the ai:* channels
 * hand out are the same for docs, sheets, slides and pdf.
 */
import { test, expect } from '@playwright/test'
import { createServer, type Server } from 'node:http'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import type { AddressInfo } from 'node:net'
import { launchShell, closeAndSaveVideo, screenshotPath } from './helpers'

interface StubCall {
  path: string
  authorization: string | undefined
  model: unknown
  stream: boolean
  /** the parsed request body, so tests can assert on fields that must be absent */
  body: Record<string, unknown>
}

/**
 * Rejects a request the way OpenAI's reasoning models do when `temperature` is
 * sent at all — the failure this feature exists to make configurable.
 */
function rejectsTemperature(body: Record<string, unknown>): boolean {
  return body.model === 'reasoning-only' && 'temperature' in body
}

/**
 * Minimal /chat/completions stub: records what it was called with and answers
 * "OK", as plain JSON or as SSE depending on the request's `stream` flag.
 */
async function startStubModelServer(): Promise<{
  baseUrl: string
  calls: StubCall[]
  stop: () => Promise<void>
}> {
  const calls: StubCall[] = []
  const server: Server = createServer((req, res) => {
    let body = ''
    req.on('data', (chunk) => (body += chunk))
    req.on('end', () => {
      let parsed: Record<string, unknown> = {}
      try {
        parsed = JSON.parse(body) as Record<string, unknown>
      } catch {
        /* a malformed body still counts as a call */
      }
      const stream = parsed.stream === true
      calls.push({
        path: req.url ?? '',
        authorization: req.headers.authorization,
        model: parsed.model,
        stream,
        body: parsed,
      })
      if (rejectsTemperature(parsed)) {
        res.writeHead(400, { 'Content-Type': 'application/json' })
        res.end(
          JSON.stringify({
            error: { message: "Unsupported value: 'temperature' does not support 0.3" },
          }),
        )
        return
      }
      if (stream) {
        res.writeHead(200, { 'Content-Type': 'text/event-stream' })
        res.write(`data: ${JSON.stringify({ choices: [{ delta: { content: 'OK' } }] })}\n\n`)
        res.write(`data: ${JSON.stringify({ choices: [{ finish_reason: 'stop' }] })}\n\n`)
        res.end('data: [DONE]\n\n')
        return
      }
      res.writeHead(200, { 'Content-Type': 'application/json' })
      res.end(
        JSON.stringify({
          choices: [{ message: { role: 'assistant', content: 'OK' }, finish_reason: 'stop' }],
        }),
      )
    })
  })
  await new Promise<void>((resolvePromise) => server.listen(0, '127.0.0.1', resolvePromise))
  const { port } = server.address() as AddressInfo
  return {
    baseUrl: `http://127.0.0.1:${port}/v1`,
    calls,
    stop: () =>
      new Promise<void>((resolvePromise) => {
        server.close(() => resolvePromise())
      }),
  }
}

test.describe('AI model settings', () => {
  test('configures a custom endpoint once and applies it to every module', async () => {
    const stub = await startStubModelServer()
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'ai-model-settings' })
    const { app, page } = launched
    try {
      // ── open the dialog from the account menu ──────────────────────
      await page.locator('.account-btn').click()
      const aiRow = page.locator('.account-menu .lang-row', { hasText: 'AI Model' })
      await expect(aiRow).toBeVisible()
      // default state: Genspark, before anything is configured
      await expect(aiRow.locator('.lang-row-current')).toHaveText('Genspark')
      await aiRow.click()

      const dialog = page.locator('.ai-modal')
      await expect(dialog).toBeVisible()
      await expect(dialog.locator('.ai-modal-scope')).toContainText('Docs, Sheets, Slides and PDF')
      await page.screenshot({ path: screenshotPath('ai-settings-genspark') })

      // ── switch to the custom endpoint and fill it in ───────────────
      await dialog.locator('.ai-mode', { hasText: 'Custom model' }).click()
      await expect(dialog.locator('.ai-fields')).toBeVisible()

      // save stays blocked until both required fields are present
      await expect(dialog.locator('.btn-primary')).toBeDisabled()
      await expect(dialog.locator('.ai-need-fields')).toBeVisible()

      const fields = dialog.locator('.ai-input')
      await fields.nth(0).fill(stub.baseUrl)
      await fields.nth(1).fill('my-local-model')
      await fields.nth(2).fill('sk-test-key')
      await expect(dialog.locator('.btn-primary')).toBeEnabled()

      // the key is masked until the reveal toggle is used
      await expect(fields.nth(2)).toHaveAttribute('type', 'password')
      await dialog.locator('.ai-key-toggle').click()
      await expect(fields.nth(2)).toHaveAttribute('type', 'text')

      // ── test connection actually reaches the endpoint ──────────────
      await dialog.locator('.btn-secondary', { hasText: 'Test connection' }).click()
      await expect(dialog.locator('.ai-test-ok')).toHaveText('Connected')
      expect(stub.calls).toHaveLength(1)
      expect(stub.calls[0].path).toBe('/v1/chat/completions')
      expect(stub.calls[0].authorization).toBe('Bearer sk-test-key')
      expect(stub.calls[0].model).toBe('my-local-model')
      await page.screenshot({ path: screenshotPath('ai-settings-custom') })

      // ── save, and confirm the menu reflects it ─────────────────────
      await dialog.locator('.btn-primary', { hasText: 'Save' }).click()
      await expect(dialog).toHaveCount(0)
      await page.locator('.account-btn').click()
      await expect(
        page
          .locator('.account-menu .lang-row', { hasText: 'AI Model' })
          .locator('.lang-row-current'),
      ).toHaveText('my-local-model')
      await page.keyboard.press('Escape')

      // ── persisted to the one settings file all modules read ────────
      const raw = await readFile(join(launched.userDataDir, 'ai-settings.json'), 'utf-8')
      expect(JSON.parse(raw)).toMatchObject({
        provider: 'custom',
        providers: {
          custom: { baseUrl: stub.baseUrl, model: 'my-local-model', apiKey: 'sk-test-key' },
        },
      })

      // ── the shared ai:* channels hand the same config to every module ──
      const resolved = await app.evaluate(async ({ ipcMain }) => {
        // invoke the real handlers the way each editor's preload does
        type Handler = (event: unknown, ...args: unknown[]) => unknown
        const invoke = (channel: string, ...args: unknown[]) =>
          (
            (ipcMain as unknown as { _invokeHandlers: Map<string, Handler> })._invokeHandlers.get(
              channel,
            ) as Handler
          )({ sender: { isDestroyed: () => false } }, ...args)
        return {
          settings: await invoke('ai:get-settings'),
          model: await invoke('ai:get-model-settings'),
        }
      })
      expect(resolved.settings).toMatchObject({ provider: 'custom' })
      expect(resolved.model).toMatchObject({ mode: 'custom', model: 'my-local-model' })

      // ── a real ai:stream turn lands on the custom endpoint ─────────
      // The request deliberately carries a stale "genspark" snapshot, the way
      // an editor tab opened before the settings change would: the file on
      // disk must win, or the change would not reach open tabs.
      const chunks = await app.evaluate(
        async ({ ipcMain }, staleSettings) => {
          type Handler = (event: unknown, ...args: unknown[]) => unknown
          const received: Array<{ type: string; text?: string; error?: string }> = []
          const sender = {
            isDestroyed: () => false,
            send: (_channel: string, chunk: { type: string; text?: string; error?: string }) =>
              received.push(chunk),
          }
          const handler = (
            ipcMain as unknown as { _invokeHandlers: Map<string, Handler> }
          )._invokeHandlers.get('ai:stream') as Handler
          await handler(
            { sender },
            {
              requestId: 'e2e-1',
              settings: staleSettings,
              system: 'You are a test.',
              messages: [{ role: 'user', text: 'ping' }],
            },
          )
          return received
        },
        { provider: 'genspark', providers: { genspark: { apiKey: '', model: 'claude-opus-4-7' } } },
      )

      expect(chunks.find((c) => c.type === 'error')).toBeUndefined()
      expect(
        chunks
          .filter((c) => c.type === 'delta')
          .map((c) => c.text)
          .join(''),
      ).toBe('OK')
      const streamCall = stub.calls.find((c) => c.stream)
      expect(streamCall).toBeDefined()
      expect(streamCall?.model).toBe('my-local-model')
      expect(streamCall?.authorization).toBe('Bearer sk-test-key')
    } finally {
      await closeAndSaveVideo(launched, 'ai-model-settings')
      await stub.stop()
    }
  })

  test('tunes temperature, max tokens and reasoning effort — including omitting temperature', async () => {
    const stub = await startStubModelServer()
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'ai-model-tuning' })
    const { app, page } = launched
    try {
      await page.locator('.account-btn').click()
      await page.locator('.account-menu .lang-row', { hasText: 'AI Model' }).click()
      const dialog = page.locator('.ai-modal')
      await dialog.locator('.ai-mode', { hasText: 'Custom model' }).click()

      const fields = dialog.locator('.ai-input')
      await fields.nth(0).fill(stub.baseUrl)
      // this model rejects any request carrying a temperature
      await fields.nth(1).fill('reasoning-only')

      const temperature = dialog.locator('.ai-tuning .ai-input').nth(0)
      const maxTokens = dialog.locator('.ai-tuning .ai-input').nth(1)
      const effort = dialog.locator('.ai-tuning select.ai-input')

      // blank means "don't send", which is what the placeholder promises
      await expect(temperature).toHaveAttribute('placeholder', 'Model default')
      await expect(temperature).toHaveValue('')

      // out-of-range input blocks saving instead of reaching the backend
      await temperature.fill('5')
      await expect(dialog.locator('.ai-field-hint.error')).toBeVisible()
      await expect(dialog.locator('.btn-primary')).toBeDisabled()
      await maxTokens.fill('0')
      await expect(dialog.locator('.btn-primary')).toBeDisabled()

      // a sent temperature is what breaks this model — prove it, then clear it
      await temperature.fill('0.3')
      await maxTokens.fill('4096')
      await effort.selectOption('high')
      await dialog.locator('.btn-secondary', { hasText: 'Test connection' }).click()
      await expect(dialog.locator('.ai-test-fail')).toBeVisible()
      await expect(dialog.locator('.ai-test-fail')).toContainText('temperature')

      await temperature.fill('')
      await dialog.locator('.btn-secondary', { hasText: 'Test connection' }).click()
      await expect(dialog.locator('.ai-test-ok')).toHaveText('Connected')

      const tested = stub.calls.at(-1)
      expect(tested?.body).not.toHaveProperty('temperature')
      expect(tested?.body.max_tokens).toBe(4096)
      expect(tested?.body.reasoning_effort).toBe('high')
      await page.screenshot({ path: screenshotPath('ai-settings-tuning') })

      await dialog.locator('.btn-primary', { hasText: 'Save' }).click()
      await expect(dialog).toHaveCount(0)

      // the same knobs apply to a real streaming turn from any editor module
      const chunks = await app.evaluate(async ({ ipcMain }) => {
        type Handler = (event: unknown, ...args: unknown[]) => unknown
        const received: Array<{ type: string; text?: string; error?: string }> = []
        const sender = {
          isDestroyed: () => false,
          send: (_c: string, chunk: { type: string; text?: string; error?: string }) =>
            received.push(chunk),
        }
        const handler = (
          ipcMain as unknown as { _invokeHandlers: Map<string, Handler> }
        )._invokeHandlers.get('ai:stream') as Handler
        await handler(
          { sender },
          {
            requestId: 'e2e-tuning',
            settings: { provider: 'genspark', providers: {} },
            system: 'You are a test.',
            // an app asking for far more than the user's ceiling must be capped
            maxTokens: 100_000,
            messages: [{ role: 'user', text: 'ping' }],
          },
        )
        return received
      })
      expect(chunks.find((c) => c.type === 'error')).toBeUndefined()

      const streamed = stub.calls.filter((c) => c.stream).at(-1)
      expect(streamed?.body).not.toHaveProperty('temperature')
      expect(streamed?.body.max_tokens).toBe(4096)
      expect(streamed?.body.reasoning_effort).toBe('high')

      // and they survive a reopen of the dialog
      await page.locator('.account-btn').click()
      await page.locator('.account-menu .lang-row', { hasText: 'AI Model' }).click()
      await expect(page.locator('.ai-tuning .ai-input').nth(0)).toHaveValue('')
      await expect(page.locator('.ai-tuning .ai-input').nth(1)).toHaveValue('4096')
      await expect(page.locator('.ai-tuning select.ai-input')).toHaveValue('high')
    } finally {
      await closeAndSaveVideo(launched, 'ai-model-tuning')
      await stub.stop()
    }
  })

  test('rejects an unreachable endpoint and keeps Genspark selected', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'ai-model-bad-endpoint' })
    const { page } = launched
    try {
      await page.locator('.account-btn').click()
      await page.locator('.account-menu .lang-row', { hasText: 'AI Model' }).click()

      const dialog = page.locator('.ai-modal')
      await dialog.locator('.ai-mode', { hasText: 'Custom model' }).click()
      const fields = dialog.locator('.ai-input')
      // port 1 is reserved and never listening
      await fields.nth(0).fill('http://127.0.0.1:1/v1')
      await fields.nth(1).fill('nope')

      await dialog.locator('.btn-secondary', { hasText: 'Test connection' }).click()
      await expect(dialog.locator('.ai-test-fail')).toBeVisible({ timeout: 30_000 })

      // a failed test does not block saving, but cancelling leaves Genspark live
      await dialog.locator('.btn-secondary', { hasText: 'Cancel' }).click()
      await expect(dialog).toHaveCount(0)
      await page.locator('.account-btn').click()
      await expect(
        page
          .locator('.account-menu .lang-row', { hasText: 'AI Model' })
          .locator('.lang-row-current'),
      ).toHaveText('Genspark')
    } finally {
      await closeAndSaveVideo(launched, 'ai-model-bad-endpoint')
    }
  })
})

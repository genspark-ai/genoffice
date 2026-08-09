/**
 * The AI Model section of the settings window, and what it actually puts on
 * the wire. Runs against a stub OpenAI-compatible server so the assertions are
 * about the request GenOffice sends, not about any live provider.
 */
import { test, expect } from '@playwright/test'
import { createServer, type Server } from 'node:http'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import type { AddressInfo } from 'node:net'
import { launchShell, closeAndSaveVideo, screenshotPath, waitForPageWithUrl } from './helpers'

interface StubCall {
  path: string
  authorization: string | undefined
  model: unknown
  stream: boolean
  body: Record<string, unknown>
}

/** rejects `temperature` the way OpenAI's reasoning models do */
function rejectsTemperature(body: Record<string, unknown>): boolean {
  return body.model === 'reasoning-only' && 'temperature' in body
}

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
  test('configures a custom endpoint in the settings window and routes every module to it', async () => {
    const stub = await startStubModelServer()
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'ai-model-settings' })
    const { app, page, userDataDir } = launched
    try {
      await page.locator('.account-btn').click()
      await page.locator('.account-menu .lang-row', { hasText: 'Settings' }).click()
      const settings = await waitForPageWithUrl(app, 'settings.html')

      // Genspark is the default until a complete custom endpoint is saved
      await expect(settings.locator('.ai-mode').first()).toHaveClass(/active/)

      await settings.locator('.ai-mode', { hasText: 'Custom model' }).click()
      const fields = settings.locator('.ai-input')
      await fields.nth(0).fill(stub.baseUrl)
      await fields.nth(1).fill('my-local-model')
      await fields.nth(2).fill('sk-test-key')

      // the key is masked until revealed
      await expect(fields.nth(2)).toHaveAttribute('type', 'password')
      await settings.locator('.ai-key-toggle').click()
      await expect(fields.nth(2)).toHaveAttribute('type', 'text')

      await settings.locator('.btn-secondary', { hasText: 'Test connection' }).click()
      await expect(settings.locator('.ai-test-ok')).toHaveText('Connected')
      expect(stub.calls[0]?.path).toBe('/v1/chat/completions')
      expect(stub.calls[0]?.authorization).toBe('Bearer sk-test-key')
      await settings.screenshot({ path: screenshotPath('settings-model-custom') })

      await settings.locator('.btn-primary', { hasText: 'Save' }).click()
      await expect(settings.locator('.settings-saved')).toBeVisible()

      const stored = JSON.parse(
        await readFile(join(userDataDir, 'ai-settings.json'), 'utf-8'),
      ) as Record<string, unknown>
      expect(stored).toMatchObject({
        provider: 'custom',
        providers: {
          custom: { baseUrl: stub.baseUrl, model: 'my-local-model', apiKey: 'sk-test-key' },
        },
      })

      // A real streaming turn, sent with a deliberately stale 'genspark'
      // snapshot the way an editor tab opened before the change would: the
      // file on disk must win, or the change would never reach open tabs.
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
            requestId: 'e2e-1',
            settings: { provider: 'genspark', providers: {} },
            system: 'You are a test.',
            messages: [{ role: 'user', text: 'ping' }],
          },
        )
        return received
      })
      expect(chunks.find((c) => c.type === 'error')).toBeUndefined()
      expect(
        chunks
          .filter((c) => c.type === 'delta')
          .map((c) => c.text)
          .join(''),
      ).toBe('OK')
      const streamCall = stub.calls.find((c) => c.stream)
      expect(streamCall?.model).toBe('my-local-model')
      expect(streamCall?.authorization).toBe('Bearer sk-test-key')
    } finally {
      await closeAndSaveVideo(launched, 'ai-model-settings')
      await stub.stop()
    }
  })

  test('omits temperature so reasoning models work, and honours the token ceiling', async () => {
    const stub = await startStubModelServer()
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'ai-model-tuning' })
    const { app } = launched
    try {
      // configure through the same handler the settings window calls
      const configure = async (temperature: number | null) =>
        app.evaluate(
          async ({ ipcMain }, [chan, cfg]) => {
            type Handler = (event: unknown, ...args: unknown[]) => unknown
            const handlers = (ipcMain as unknown as { _invokeHandlers: Map<string, Handler> })
              ._invokeHandlers
            return handlers.get(chan as string)!({}, cfg)
          },
          [
            'ai:set-model-settings',
            {
              mode: 'custom',
              baseUrl: stub.baseUrl,
              model: 'reasoning-only',
              apiKey: '',
              temperature,
              maxTokens: 4096,
              reasoningEffort: 'high',
              tavilyApiKey: '',
              proxyUrl: '',
            },
          ] as [string, unknown],
        )

      const testProvider = async () =>
        app.evaluate(
          async ({ ipcMain }, cfg) => {
            type Handler = (event: unknown, ...args: unknown[]) => unknown
            const handlers = (ipcMain as unknown as { _invokeHandlers: Map<string, Handler> })
              ._invokeHandlers
            return (await handlers.get('ai:test-provider')!({}, cfg)) as {
              ok: boolean
              error?: string
            }
          },
          {
            baseUrl: stub.baseUrl,
            model: 'reasoning-only',
            apiKey: '',
            temperature: 0.3,
            maxTokens: 4096,
            reasoningEffort: 'high',
          },
        )

      // sending a temperature is exactly what this model rejects
      await configure(0.3)
      const withTemp = await testProvider()
      expect(withTemp.ok).toBe(false)
      expect(withTemp.error).toContain('temperature')

      // blank means "omit the field", which is what makes it work
      await configure(null)
      const omitted = await app.evaluate(async ({ ipcMain }) => {
        type Handler = (event: unknown, ...args: unknown[]) => unknown
        const handlers = (ipcMain as unknown as { _invokeHandlers: Map<string, Handler> })
          ._invokeHandlers
        return (await handlers.get('ai:test-provider')!(
          {},
          {
            baseUrl: '',
            model: '',
            apiKey: '',
            temperature: null,
            maxTokens: null,
            reasoningEffort: null,
          },
        )) as { ok: boolean }
      })
      // an empty draft is rejected before it reaches the network
      expect(omitted.ok).toBe(false)

      const chunks = await app.evaluate(async ({ ipcMain }) => {
        type Handler = (event: unknown, ...args: unknown[]) => unknown
        const received: Array<{ type: string; error?: string }> = []
        const sender = {
          isDestroyed: () => false,
          send: (_c: string, ch: never) => received.push(ch),
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
    } finally {
      await closeAndSaveVideo(launched, 'ai-model-tuning')
      await stub.stop()
    }
  })
})

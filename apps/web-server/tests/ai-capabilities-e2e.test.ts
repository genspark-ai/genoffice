/**
 * E2E test for home:ai-capabilities.
 *
 * Boots the full web-server bundle on a random port and verifies the
 * AI capability report correctly distinguishes:
 *   1. configured (keyed provider / genspark login) vs available
 *   2. unkeyed DuckDuckGo fallback for web/image search
 *   3. no fallback for image generation / media analysis
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { spawn, type ChildProcess } from 'node:child_process'
import { mkdtempSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

interface IpcResult<T = unknown> {
  ok: boolean
  result: T
}

async function ipc<T = unknown>(base: string, channel: string, args: unknown[] = []): Promise<IpcResult<T>> {
  const res = await fetch(`${base}/api/ipc/${channel}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ args }),
  })
  return (await res.json()) as IpcResult<T>
}

async function waitForHealth(base: string, timeoutMs = 30_000): Promise<void> {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    try {
      const res = await fetch(`${base}/health`)
      if (res.ok) return
    } catch {
      /* keep polling */
    }
    await new Promise((r) => setTimeout(r, 250))
  }
  throw new Error(`web-server did not become healthy within ${timeoutMs}ms`)
}

interface CapabilityEntry {
  available: boolean
  via: string
  fallback?: string
  configured: boolean
  note?: string
}

interface CapabilitiesReport {
  ok: boolean
  capabilities: {
    search: CapabilityEntry
    image_search: CapabilityEntry
    image_generation: CapabilityEntry
    media_analysis: CapabilityEntry
  }
  provider: string
  gskToolsEnabled: boolean
}

describe('home:ai-capabilities E2E', () => {
  let server: ChildProcess | undefined
  let base: string
  let dataDir: string
  let port: number

  beforeAll(async () => {
    dataDir = mkdtempSync(join(tmpdir(), 'genoffice-cap-'))
    port = 20000 + Math.floor(Math.random() * 9000)
    base = `http://127.0.0.1:${port}`

    // Pre-seed an ai-settings file: Serper key for search, no media key.
    writeFileSync(
      join(dataDir, 'ai-settings.json'),
      JSON.stringify({
        provider: 'anthropic',
        providers: {
          anthropic: { apiKey: '' },
        },
        search: {
          provider: 'serper',
          providers: { serper: { apiKey: 'k' }, tavily: { apiKey: '' } },
        },
        gskToolsEnabled: false,
      }),
    )

    const bundle = join(process.cwd(), 'dist/bundle/index.js')
    server = spawn('node', [bundle], {
      env: {
        ...process.env,
        GENOFFICE_DATA_DIR: dataDir,
        PORT: String(port),
      },
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    await waitForHealth(base)
  }, 30_000)

  afterAll(() => {
    if (server && !server.killed) {
      server.kill('SIGTERM')
    }
    if (dataDir) {
      try {
        const { rmSync } = require('node:fs') as typeof import('node:fs')
        rmSync(dataDir, { recursive: true, force: true })
      } catch { /* ignore */ }
    }
  })

  it('reports Serper as configured with DDG fallback for web + image search', async () => {
    const r = await ipc<CapabilitiesReport>(base, 'home:ai-capabilities', [])
    expect(r.ok).toBe(true)
    const caps = r.result.capabilities
    expect(caps.search.configured).toBe(true)
    expect(caps.search.available).toBe(true)
    expect(caps.search.via).toBe('serper')
    expect(caps.search.fallback).toBe('duckduckgo')
    expect(caps.image_search.configured).toBe(true)
    expect(caps.image_search.available).toBe(true)
    expect(caps.image_search.via).toBe('serper')
    expect(caps.image_search.fallback).toBe('duckduckgo')
  })

  it('reports image generation + media analysis as unavailable with no media key', async () => {
    const r = await ipc<CapabilitiesReport>(base, 'home:ai-capabilities', [])
    expect(r.ok).toBe(true)
    const caps = r.result.capabilities
    expect(caps.image_generation.configured).toBe(false)
    expect(caps.image_generation.available).toBe(false)
    expect(caps.media_analysis.configured).toBe(false)
    expect(caps.media_analysis.available).toBe(false)
  })

  it('honors gskToolsEnabled=false by making generation + analysis unavailable', async () => {
    // The seeded settings set gskToolsEnabled=false, so even the genspark
    // path is closed off for media — confirm the report reflects that.
    const r = await ipc<CapabilitiesReport>(base, 'home:ai-capabilities', [])
    expect(r.result.gskToolsEnabled).toBe(false)
    expect(r.result.capabilities.image_generation.via).toBe('none')
    expect(r.result.capabilities.media_analysis.via).toBe('none')
  })
})

import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { run, tempDir } from './helpers'

// hasGskAuth reads process.env, not the command context: isolate the login state per test
const saved: Record<string, string | undefined> = {}
beforeEach(() => {
  for (const k of ['GENOFFICE_AUTH_DIR', 'AI_SEARCH_DISABLE_GSK']) saved[k] = process.env[k]
  process.env.GENOFFICE_AUTH_DIR = join(tempDir(), 'no-auth')
  process.env.AI_SEARCH_DISABLE_GSK = '1'
})
afterEach(() => {
  for (const [k, v] of Object.entries(saved)) {
    if (v === undefined) delete process.env[k]
    else process.env[k] = v
  }
})

// a real settings file always carries the chat provider block; without it every section resets to defaults
function settingsFile(dir: string, settings: Record<string, unknown>): string {
  const path = join(dir, 'ai-settings.json')
  writeFileSync(path, JSON.stringify({ provider: 'anthropic', providers: {}, ...settings }))
  return path
}

describe('genoffice capabilities', () => {
  it('reports web/image search available via DDG fallback even when signed out with default settings', async () => {
    const dir = tempDir()
    const r = await run(['capabilities', '--json'], {
      env: {
        ...process.env,
        GENOFFICE_AI_SETTINGS: join(dir, 'missing.json'),
        GENOFFICE_APP_BIN: '',
      },
    })
    expect(r.code).toBe(0)
    const d = r.json().detail
    // Web/image search stay usable through the DuckDuckGo fallback.
    expect(d.search.configured).toBe(false)
    expect(d.search.available).toBe(true)
    expect(d.search.via).toBe('duckduckgo')
    expect(d.image_search.configured).toBe(false)
    expect(d.image_search.available).toBe(true)
    expect(d.image_search.via).toBe('duckduckgo')
    // Generation + analysis have no fallback, so they remain unavailable.
    expect(d.image_generation.available).toBe(false)
    expect(d.media_analysis.available).toBe(false)
  })

  it('counts a Serper key as search + image search and a BYOK image model as generation', async () => {
    const dir = tempDir()
    mkdirSync(join(dir, 'bin'))
    const settings = settingsFile(dir, {
      search: {
        provider: 'serper',
        providers: { serper: { apiKey: 'k' }, tavily: { apiKey: '' } },
      },
      media: {
        imageProvider: 'openai',
        providers: { openai: { apiKey: 'sk', imageModel: 'gpt-image-1' } },
      },
    })
    const r = await run(['capabilities', '--json'], {
      env: {
        ...process.env,
        GENOFFICE_AI_SETTINGS: settings,
        GENOFFICE_APP_BIN: join(dir, 'bin', 'app'),
      },
    })
    expect(r.code).toBe(0)
    const d = r.json().detail
    expect(d.search.configured).toBe(true)
    expect(d.search.available).toBe(true)
    expect(d.search.via).toBe('serper')
    expect(d.search.fallback).toBe('duckduckgo')
    expect(d.image_search.configured).toBe(true)
    expect(d.image_search.via).toBe('serper')
    expect(d.image_generation.configured).toBe(true)
    expect(d.image_generation.available).toBe(true)
    expect(d.image_generation.via).toBe('openai')
    expect(d.media_analysis.available).toBe(false)
    expect(d.app.available).toBe(true)
    expect(r.json().summary).toContain('image_generation')
  })

  it('Tavily gives web search (with DDG fallback) but no configured image search', async () => {
    const dir = tempDir()
    const settings = settingsFile(dir, {
      search: {
        provider: 'tavily',
        providers: { serper: { apiKey: '' }, tavily: { apiKey: 't' } },
      },
    })
    const r = await run(['capabilities', '--json'], {
      env: { ...process.env, GENOFFICE_AI_SETTINGS: settings },
    })
    const d = r.json().detail
    expect(d.search.configured).toBe(true)
    expect(d.search.via).toBe('tavily')
    // Image search is NOT configured (Tavily does not serve images), but
    // the DDG fallback still makes `available` true.
    expect(d.image_search.configured).toBe(false)
    expect(d.image_search.available).toBe(true)
    expect(d.image_search.via).toBe('duckduckgo')
  })
})

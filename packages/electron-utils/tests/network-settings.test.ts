import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { gskProxyUrl, setGskProxyUrl } from '@genoffice/ai-search'

import { bootstrapNetworkSettings, readNetworkSettings } from '../src/network-settings'

let userData: string

beforeEach(() => {
  userData = mkdtempSync(join(tmpdir(), 'genoffice-network-'))
})

afterEach(() => {
  rmSync(userData, { recursive: true, force: true })
  setGskProxyUrl('')
})

function writeSettings(value: unknown): void {
  writeFileSync(join(userData, 'ai-settings.json'), JSON.stringify(value), 'utf-8')
}

describe('readNetworkSettings', () => {
  it('returns empty fields when the settings file does not exist', () => {
    expect(readNetworkSettings(userData)).toEqual({ proxyUrl: '', tavilyApiKey: '' })
  })

  it('returns empty fields when the settings file is corrupt', () => {
    writeFileSync(join(userData, 'ai-settings.json'), '{ not json', 'utf-8')
    expect(readNetworkSettings(userData)).toEqual({ proxyUrl: '', tavilyApiKey: '' })
  })

  it('reads the network half and ignores the provider half', () => {
    writeSettings({
      provider: 'custom',
      providers: { custom: { baseUrl: 'http://localhost:11434/v1', model: 'qwen3', apiKey: '' } },
      proxyUrl: 'socks5://127.0.0.1:7897',
      tavilyApiKey: 'tvly-abc',
    })
    expect(readNetworkSettings(userData)).toEqual({
      proxyUrl: 'socks5://127.0.0.1:7897',
      tavilyApiKey: 'tvly-abc',
    })
  })
})

// applyProxy remembers the last url it wired up for the lifetime of the process,
// so a no-op save does not churn undici and Chromium. That memo is module state
// these tests share: each one below picks a proxy no earlier test used, so the
// call it is asserting on is never the one applyProxy skips.
describe('bootstrapNetworkSettings', () => {
  it('reports an explicit proxy so the caller skips env/system detection', () => {
    writeSettings({ proxyUrl: 'http://127.0.0.1:7001' })
    expect(bootstrapNetworkSettings(userData)).toBe(true)
  })

  it('reports no explicit proxy when the field is blank, so env/system still apply', () => {
    writeSettings({ proxyUrl: '', tavilyApiKey: 'tvly-abc' })
    expect(bootstrapNetworkSettings(userData)).toBe(false)
  })

  it('reports no explicit proxy for an unusable value rather than wiring it up', () => {
    writeSettings({ proxyUrl: 'not a proxy' })
    expect(bootstrapNetworkSettings(userData)).toBe(false)
  })

  it('forwards the proxy to gsk CLI children — the whole point for PPT generation', () => {
    writeSettings({ proxyUrl: '127.0.0.1:7002' })
    expect(bootstrapNetworkSettings(userData)).toBe(true)
    // bare host:port is normalized to http:// before it reaches the child env
    expect(gskProxyUrl()).toBe('http://127.0.0.1:7002')
  })

  it('clears a previously forwarded proxy when the user empties the field', () => {
    writeSettings({ proxyUrl: 'http://127.0.0.1:7003' })
    bootstrapNetworkSettings(userData)
    expect(gskProxyUrl()).toBe('http://127.0.0.1:7003')

    writeSettings({ proxyUrl: '' })
    expect(bootstrapNetworkSettings(userData)).toBe(false)
    expect(gskProxyUrl()).toBe('')
  })

  it('leaves the wiring alone when the same proxy is saved again', () => {
    writeSettings({ proxyUrl: 'http://127.0.0.1:7004' })
    bootstrapNetworkSettings(userData)
    expect(gskProxyUrl()).toBe('http://127.0.0.1:7004')

    // clear the forwarded value by hand: a re-bootstrap of the same url must
    // not push it again, which is what proves the memo is doing its job
    setGskProxyUrl('')
    expect(bootstrapNetworkSettings(userData)).toBe(true)
    expect(gskProxyUrl()).toBe('')
  })
})

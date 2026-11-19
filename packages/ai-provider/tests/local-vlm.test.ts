import { EventEmitter } from 'node:events'
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  analyzeMediaWithProvider,
  canAttemptChat,
  chatForProvider,
  defaultAiMediaSettings,
  defaultAiSettings,
  getProviderAdapter,
  isKeylessProvider,
  isLocalVlmProvider,
  isValidGgufFile,
  LOCAL_VLM_DEFAULT_MODEL,
  LOCAL_VLM_PROVIDER_ID,
  localVlmSettingsFrom,
  LocalVlmManager,
  mergeLocalVlmSettings,
  registerLocalVlmRuntime,
  selectLlamaReleaseAsset,
  shouldFallbackFromAgyToLocalVlm,
  streamForProvider,
  testMediaProvider,
  type LocalVlmRuntimeBridge,
} from '../src/index'
import { setAgyUsable } from '../src/agy-default'

const tempDirs: string[] = []
function makeTempDir(prefix = 'genoffice-localvlm-test-'): string {
  const dir = mkdtempSync(join(tmpdir(), prefix))
  tempDirs.push(dir)
  return dir
}

function makeFakeGguf(filePath: string): void {
  mkdirSync(join(filePath, '..'), { recursive: true })
  const buf = Buffer.alloc(64)
  Buffer.from('GGUF', 'ascii').copy(buf, 0)
  writeFileSync(filePath, buf)
}

function makeFakeRuntimeExe(filePath: string): void {
  mkdirSync(join(filePath, '..'), { recursive: true })
  writeFileSync(filePath, '#!/bin/sh\necho ok\n')
}

afterEach(() => {
  vi.unstubAllGlobals()
  registerLocalVlmRuntime(null)
  setAgyUsable(null)
  while (tempDirs.length) {
    const dir = tempDirs.pop()!
    rmSync(dir, { recursive: true, force: true })
  }
})

describe('local-vlm metadata and settings', () => {
  it('merges and clamps local VLM settings cleanly', () => {
    const merged = mergeLocalVlmSettings(undefined, {
      enabled: false,
      autoFallbackFromAgy: false,
      ctxSize: 999999,
      threads: -5,
      gpuLayers: 32,
      modelPath: '  /tmp/model.gguf  ',
    })
    expect(merged).toEqual({
      enabled: false,
      autoFallbackFromAgy: false,
      modelPath: '/tmp/model.gguf',
      mmprojPath: '',
      runtimePath: '',
      ctxSize: 32768,
      threads: 0,
      gpuLayers: 32,
    })
    expect(localVlmSettingsFrom({})).toEqual(mergeLocalVlmSettings(undefined, {}))
  })

  it('defaults chat and analysis to local-vlm only when agy is unusable and local-vlm is usable', () => {
    const whenNothingReady = defaultAiSettings(undefined, false, false)
    expect(whenNothingReady.provider).toBe('gemini')
    expect(whenNothingReady.media?.analysisProvider).toBe('gemini')

    const whenLocalReady = defaultAiSettings(undefined, false, true)
    expect(whenLocalReady.provider).toBe('local-vlm')
    expect(whenLocalReady.providers['local-vlm'].model).toBe(LOCAL_VLM_DEFAULT_MODEL)
    expect(whenLocalReady.media?.analysisProvider).toBe('local-vlm')
    expect(whenLocalReady.media?.imageProvider).toBe('gemini')

    const whenAgyReady = defaultAiSettings(undefined, true, true)
    expect(whenAgyReady.provider).toBe('agy')
    expect(defaultAiMediaSettings(false, true).analysisProvider).toBe('local-vlm')
  })

  it('registers local-vlm as keyless, vision-capable, and ready to attempt chat', () => {
    expect(isLocalVlmProvider(LOCAL_VLM_PROVIDER_ID)).toBe(true)
    expect(isKeylessProvider(LOCAL_VLM_PROVIDER_ID)).toBe(true)
    const adapter = getProviderAdapter('local-vlm')
    expect(adapter.capabilities.vision).toBe(true)
    expect(adapter.capabilities.tools).toBe(true)
    expect(adapter.resolveEndpoint({ apiKey: '', model: '' }).protocol).toBe('local-vlm')
    expect(canAttemptChat(defaultAiSettings(undefined, false, true))).toBe(true)
  })
})

describe('selectLlamaReleaseAsset & GGUF validation', () => {
  it('picks matching CPU assets across Windows, macOS, and Linux', () => {
    const assets = [
      {
        name: 'llama-b12252-bin-win-cpu-x64.zip',
        browser_download_url: 'https://example/win-x64.zip',
      },
      {
        name: 'llama-b12252-bin-win-cpu-arm64.zip',
        browser_download_url: 'https://example/win-arm64.zip',
      },
      {
        name: 'llama-b12252-bin-macos-arm64.tar.gz',
        browser_download_url: 'https://example/mac-arm64.tar.gz',
      },
      {
        name: 'llama-b12252-bin-ubuntu-x64.tar.gz',
        browser_download_url: 'https://example/linux-x64.tar.gz',
      },
    ]
    expect(selectLlamaReleaseAsset(assets, 'win32', 'x64')?.name).toBe(
      'llama-b12252-bin-win-cpu-x64.zip',
    )
    expect(selectLlamaReleaseAsset(assets, 'win32', 'arm64')?.name).toBe(
      'llama-b12252-bin-win-cpu-arm64.zip',
    )
    expect(selectLlamaReleaseAsset(assets, 'darwin', 'arm64')?.name).toBe(
      'llama-b12252-bin-macos-arm64.tar.gz',
    )
    expect(selectLlamaReleaseAsset(assets, 'linux', 'x64')?.name).toBe(
      'llama-b12252-bin-ubuntu-x64.tar.gz',
    )
  })

  it('validates GGUF magic header bytes', () => {
    const dir = makeTempDir()
    const valid = join(dir, 'ok.gguf')
    const invalid = join(dir, 'bad.gguf')
    makeFakeGguf(valid)
    writeFileSync(invalid, Buffer.from('NOT_GGUF_HEADER_DATA_1234'))
    expect(isValidGgufFile(valid).valid).toBe(true)
    expect(isValidGgufFile(invalid).valid).toBe(false)
    expect(isValidGgufFile(join(dir, 'missing.gguf')).valid).toBe(false)
  })
})

describe('LocalVlmManager lifecycle & download', () => {
  it('starts llama-server once for concurrent callers, tests connection, and stops cleanly', async () => {
    const dir = makeTempDir()
    const mgr = new LocalVlmManager({
      userDataPath: dir,
      platform: 'linux',
      allocatePort: async () => 45678,
      freeRamMB: () => 4096,
      fetchFn: (async () => new Response('{"status":"ok"}', { status: 200 })) as typeof fetch,
      spawnFn: (() => {
        const proc = new EventEmitter() as EventEmitter & {
          pid: number
          exitCode: number | null
          signalCode: string | null
          stderr: EventEmitter
          kill: () => boolean
        }
        proc.pid = 99999
        proc.exitCode = null
        proc.signalCode = null
        proc.stderr = new EventEmitter()
        proc.kill = () => {
          proc.exitCode = 0
          proc.emit('exit', 0, null)
          return true
        }
        return proc
      }) as unknown as LocalVlmManager extends never
        ? never
        : NonNullable<ConstructorParameters<typeof LocalVlmManager>[0]['spawnFn']>,
    })

    makeFakeGguf(mgr.defaultModelPath())
    makeFakeGguf(mgr.defaultMmprojPath())
    makeFakeRuntimeExe(join(mgr.runtimeDir(), 'llama-server'))

    expect(mgr.isRunnable()).toBe(true)
    expect(mgr.shouldFallbackFromAgy()).toBe(true)

    const [a, b] = await Promise.all([mgr.ensureRunning(), mgr.ensureRunning()])
    expect(a.baseUrl).toBe('http://127.0.0.1:45678/v1')
    expect(b.model).toBe(LOCAL_VLM_DEFAULT_MODEL)
    expect(mgr.status().running).toBe(true)

    const testRes = await mgr.testConnection()
    expect(testRes.ok).toBe(true)

    await mgr.stop()
    expect(mgr.status().running).toBe(false)
  })

  it('downloads runtime + GGUF model + mmproj atomically from mocked HTTP responses', async () => {
    const dir = makeTempDir()
    const ggufPayload = Buffer.alloc(64)
    Buffer.from('GGUF', 'ascii').copy(ggufPayload, 0)

    const fetchMock = (async (input: RequestInfo | URL) => {
      const url = String(input)
      if (url.includes('releases/latest')) {
        return new Response(
          JSON.stringify({
            assets: [
              {
                name: 'llama-b12252-bin-ubuntu-x64.tar.gz',
                browser_download_url: 'https://example.invalid/llama.tar.gz',
              },
            ],
          }),
          { status: 200 },
        )
      }
      return new Response(ggufPayload, {
        status: 200,
        headers: { 'content-length': String(ggufPayload.byteLength) },
      })
    }) as typeof fetch

    const mgr = new LocalVlmManager({
      userDataPath: dir,
      platform: 'linux',
      arch: 'x64',
      fetchFn: fetchMock,
      extractArchive: async (_archive, dest) => {
        makeFakeRuntimeExe(join(dest, 'llama-server'))
      },
    })

    const status = await mgr.downloadAll()
    expect(status.installed).toBe(true)
    expect(status.runtimeInstalled).toBe(true)
    expect(mgr.isRunnable()).toBe(true)
  })
})

describe('agy fallback and OpenAI-compatible routing', () => {
  it('routes direct local-vlm and agy fallback requests to local server when agy is missing', async () => {
    const requests: Array<{ url: string; body: Record<string, unknown> }> = []
    vi.stubGlobal('fetch', (async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input)
      if (url.endsWith('/models')) {
        return new Response(JSON.stringify({ data: [{ id: LOCAL_VLM_DEFAULT_MODEL }] }), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        })
      }
      const parsed = JSON.parse(String(init?.body ?? '{}')) as Record<string, unknown>
      requests.push({ url, body: parsed })
      if (parsed.stream) {
        const sse = [
          `data: ${JSON.stringify({ choices: [{ delta: { content: 'Local ' }, finish_reason: null }] })}\n\n`,
          `data: ${JSON.stringify({ choices: [{ delta: { content: 'LFM3-VL' }, finish_reason: 'stop' }] })}\n\n`,
          'data: [DONE]\n\n',
        ].join('')
        return new Response(sse, {
          status: 200,
          headers: { 'Content-Type': 'text/event-stream' },
        })
      }
      return new Response(
        JSON.stringify({
          choices: [{ message: { content: 'Local vision answer' }, finish_reason: 'stop' }],
        }),
        { status: 200, headers: { 'Content-Type': 'application/json' } },
      )
    }) as typeof fetch)

    const baseUrl = 'http://127.0.0.1:45679/v1'
    const bridge: LocalVlmRuntimeBridge = {
      status: () => ({
        installed: true,
        modelInstalled: true,
        mmprojInstalled: true,
        runtimeInstalled: true,
        running: true,
        starting: false,
        port: 45679,
        modelPath: '/tmp/m.gguf',
        mmprojPath: '/tmp/mm.gguf',
        availableRamMB: 4096,
        ramSufficient: true,
        settings: mergeLocalVlmSettings(undefined, {}),
      }),
      isRunnable: () => true,
      shouldFallbackFromAgy: () => true,
      ensureRunning: async () => ({ baseUrl, model: LOCAL_VLM_DEFAULT_MODEL }),
      testConnection: async () => ({ ok: true, latencyMs: 5 }),
      stop: async () => undefined,
    }
    registerLocalVlmRuntime(bridge)

    // Force agy missing check to return true quickly via explicit deps
    expect(
      await shouldFallbackFromAgyToLocalVlm(undefined, {
        resolveCli: async () => {
          throw new Error('not found')
        },
      }),
    ).toBe(true)

    // Direct local-vlm one-shot chat
    const chatRes = await chatForProvider('local-vlm', { apiKey: '', model: '' }, 'sys', 'hi')
    expect(chatRes).toEqual({ ok: true, content: 'Local vision answer' })

    // Direct local-vlm stream
    let streamed = ''
    await streamForProvider(
      'local-vlm',
      { apiKey: '', model: '' },
      'sys',
      [{ role: 'user', text: 'hi' }],
      [],
      256,
      {
        signal: new AbortController().signal,
        onDelta: (d) => (streamed += d),
        onToolCall: () => undefined,
      },
    )
    expect(streamed).toBe('Local LFM3-VL')

    // Media analysis via local-vlm
    const mediaText = await analyzeMediaWithProvider(
      'local-vlm',
      { apiKey: '', imageModel: '', analysisModel: LOCAL_VLM_DEFAULT_MODEL },
      {
        requirements: 'Read text in image',
        media: [{ bytes: Buffer.from([1, 2, 3]), mime: 'image/png', name: 'scan.png' }],
      },
    )
    expect(mediaText).toBe('Local vision answer')

    // testMediaProvider for local-vlm
    expect(
      await testMediaProvider('local-vlm', {
        apiKey: '',
        imageModel: '',
        analysisModel: LOCAL_VLM_DEFAULT_MODEL,
      }),
    ).toEqual({ ok: true, latencyMs: 5 })

    // agy fallback when agy is missing
    setAgyUsable(false)
    const agyFallbackChat = await chatForProvider(
      'agy',
      { apiKey: '', model: '', cliPath: '/nonexistent/agy-binary-missing' },
      'sys',
      'hello',
    )
    expect(agyFallbackChat).toEqual({ ok: true, content: 'Local vision answer' })
    expect(requests.length).toBeGreaterThanOrEqual(4)
  })

  it('does not fall back when agy is usable', async () => {
    const dir = makeTempDir()
    const fakeCli = join(dir, 'agy')
    writeFileSync(fakeCli, '')
    registerLocalVlmRuntime({
      status: vi.fn(),
      isRunnable: () => true,
      shouldFallbackFromAgy: () => true,
      ensureRunning: vi.fn(),
      testConnection: vi.fn(),
      stop: vi.fn(),
    })
    const fallback = await shouldFallbackFromAgyToLocalVlm(fakeCli, {
      resolveCli: async () => fakeCli,
      agyUsable: async () => true,
    })
    expect(fallback).toBe(false)
  })

  it('intercepts XML <tool_call><function=read_blocks></function></tool_call> and emits onToolCall without leaking text', async () => {
    vi.stubGlobal('fetch', (async (_input: RequestInfo | URL, init?: RequestInit) => {
      const parsed = JSON.parse(String(init?.body ?? '{}')) as Record<string, unknown>
      if (parsed.stream) {
        const sse = [
          `data: ${JSON.stringify({ choices: [{ delta: { content: '<tool_call>\n<function=read_blocks>\n' }, finish_reason: null }] })}\n\n`,
          `data: ${JSON.stringify({ choices: [{ delta: { content: '</function>\n</tool_call>' }, finish_reason: null }] })}\n\n`,
          `data: ${JSON.stringify({ choices: [{ delta: {}, finish_reason: 'stop' }] })}\n\n`,
          'data: [DONE]\n\n',
        ].join('')
        return new Response(sse, { status: 200, headers: { 'Content-Type': 'text/event-stream' } })
      }
      return new Response('{}', { status: 200 })
    }) as typeof fetch)

    const baseUrl = 'http://127.0.0.1:45679/v1'
    registerLocalVlmRuntime({
      status: vi.fn() as never,
      isRunnable: () => true,
      shouldFallbackFromAgy: () => true,
      ensureRunning: async () => ({ baseUrl, model: LOCAL_VLM_DEFAULT_MODEL }),
      testConnection: async () => ({ ok: true }),
      stop: async () => undefined,
    })

    let textDelta = ''
    const toolCalls: unknown[] = []
    let finalStopReason: string | undefined

    await streamForProvider(
      'local-vlm',
      { apiKey: '', model: '' },
      'You are a document editing assistant',
      [{ role: 'user', text: 'can you wwiring on file' }],
      [
        {
          name: 'read_blocks',
          description: 'Read document blocks',
          inputSchema: { type: 'object', properties: {} },
        },
      ],
      512,
      {
        signal: new AbortController().signal,
        onDelta: (d) => (textDelta += d),
        onToolCall: (tc) => toolCalls.push(tc),
        onStopReason: (r) => (finalStopReason = r),
      },
    )

    expect(textDelta).toBe('')
    expect(toolCalls.length).toBe(1)
    expect(toolCalls[0]).toMatchObject({
      name: 'read_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 0 },
    })
    expect(finalStopReason).toBe('tool_calls')
  })

  it('normalizes direct formatting op like setFont into apply_ops with document target', async () => {
    const baseUrl = 'http://127.0.0.1:45679/v1'
    registerLocalVlmRuntime({
      status: vi.fn() as never,
      isRunnable: () => true,
      shouldFallbackFromAgy: () => true,
      ensureRunning: async () => ({ baseUrl, model: LOCAL_VLM_DEFAULT_MODEL }),
      testConnection: async () => ({ ok: true }),
      stop: async () => undefined,
    })

    vi.stubGlobal('fetch', (async (_input: RequestInfo | URL, init?: RequestInit) => {
      const parsed = JSON.parse(String(init?.body ?? '{}')) as Record<string, unknown>
      if (parsed.stream) {
        const sse = [
          `data: ${JSON.stringify({ choices: [{ delta: { content: '<tool_call>\n<function=setFont>\n<arguments>\n{"fontSize":13,"fontFamily":"Times New Roman"}\n</arguments>\n</function>\n</tool_call>' }, finish_reason: null }] })}\n\n`,
          `data: ${JSON.stringify({ choices: [{ delta: {}, finish_reason: 'stop' }] })}\n\n`,
          'data: [DONE]\n\n',
        ].join('')
        return new Response(sse, { status: 200, headers: { 'Content-Type': 'text/event-stream' } })
      }
      return new Response('{}', { status: 200 })
    }) as typeof fetch)

    let textDelta = ''
    const toolCalls: any[] = []
    let finalStopReason: string | undefined

    await streamForProvider(
      'local-vlm',
      { apiKey: '', model: '' },
      'You are a document editing assistant',
      [{ role: 'user', text: 'định dạng lại cỡ chữ 13 font times new roman cho tôi' }],
      [
        {
          name: 'apply_ops',
          description: 'Apply ops',
          inputSchema: { type: 'object', properties: {} },
        },
      ],
      512,
      {
        signal: new AbortController().signal,
        onDelta: (d) => (textDelta += d),
        onToolCall: (tc) => toolCalls.push(tc),
        onStopReason: (r) => (finalStopReason = r),
      },
    )

    expect(textDelta).toBe('')
    expect(toolCalls.length).toBe(1)
    expect(toolCalls[0]).toMatchObject({
      name: 'apply_ops',
      input: {
        ops: [
          {
            op: 'setFont',
            fontSize: 13,
            fontFamily: 'Times New Roman',
            target: { scope: 'document' },
          },
        ],
      },
    })
    expect(finalStopReason).toBe('tool_calls')
  })

  it('blocks deleteBlocks without explicit target — no onToolCall, no data loss', async () => {
    // Blocker 2: destructive ops without target must be silently dropped (not auto-scoped)
    const baseUrl = 'http://127.0.0.1:45679/v1'
    registerLocalVlmRuntime({
      status: vi.fn() as never,
      isRunnable: () => true,
      shouldFallbackFromAgy: () => true,
      ensureRunning: async () => ({ baseUrl, model: LOCAL_VLM_DEFAULT_MODEL }),
      testConnection: async () => ({ ok: true }),
      stop: async () => undefined,
    })

    vi.stubGlobal('fetch', (async (_input: RequestInfo | URL, init?: RequestInit) => {
      const parsed = JSON.parse(String(init?.body ?? '{}')) as Record<string, unknown>
      if (parsed.stream) {
        // Model emits deleteBlocks with no target — must be dropped
        const sse = [
          `data: ${JSON.stringify({ choices: [{ delta: { content: '<tool_call>\n<function=deleteBlocks>\n<arguments>\n{}\n</arguments>\n</function>\n</tool_call>' }, finish_reason: null }] })}\n\n`,
          `data: ${JSON.stringify({ choices: [{ delta: {}, finish_reason: 'stop' }] })}\n\n`,
          'data: [DONE]\n\n',
        ].join('')
        return new Response(sse, { status: 200, headers: { 'Content-Type': 'text/event-stream' } })
      }
      return new Response('{}', { status: 200 })
    }) as typeof fetch)

    const toolCalls: unknown[] = []
    await streamForProvider(
      'local-vlm',
      { apiKey: '', model: '' },
      'You are a document editing assistant',
      [{ role: 'user', text: 'delete everything' }],
      [{ name: 'apply_ops', description: 'Apply ops', inputSchema: { type: 'object', properties: {} } }],
      512,
      {
        signal: new AbortController().signal,
        onDelta: () => undefined,
        onToolCall: (tc) => toolCalls.push(tc),
      },
    )

    // deleteBlocks without target must be blocked — no tool call must fire
    expect(toolCalls.length).toBe(0)
  })

  it('allows deleteBlocks when model provides an explicit blockIndexes target', async () => {
    // Blocker 2 (positive case): a proper target is accepted
    const baseUrl = 'http://127.0.0.1:45679/v1'
    registerLocalVlmRuntime({
      status: vi.fn() as never,
      isRunnable: () => true,
      shouldFallbackFromAgy: () => true,
      ensureRunning: async () => ({ baseUrl, model: LOCAL_VLM_DEFAULT_MODEL }),
      testConnection: async () => ({ ok: true }),
      stop: async () => undefined,
    })

    vi.stubGlobal('fetch', (async (_input: RequestInfo | URL, init?: RequestInit) => {
      const parsed = JSON.parse(String(init?.body ?? '{}')) as Record<string, unknown>
      if (parsed.stream) {
        const sse = [
          `data: ${JSON.stringify({ choices: [{ delta: { content: '<tool_call>\n<function=deleteBlocks>\n<arguments>\n{"target":{"blockIndexes":[2,3]}}\n</arguments>\n</function>\n</tool_call>' }, finish_reason: null }] })}\n\n`,
          `data: ${JSON.stringify({ choices: [{ delta: {}, finish_reason: 'stop' }] })}\n\n`,
          'data: [DONE]\n\n',
        ].join('')
        return new Response(sse, { status: 200, headers: { 'Content-Type': 'text/event-stream' } })
      }
      return new Response('{}', { status: 200 })
    }) as typeof fetch)

    const toolCalls: any[] = []
    await streamForProvider(
      'local-vlm',
      { apiKey: '', model: '' },
      'You are a document editing assistant',
      [{ role: 'user', text: 'delete blocks 2 and 3' }],
      [{ name: 'apply_ops', description: 'Apply ops', inputSchema: { type: 'object', properties: {} } }],
      512,
      {
        signal: new AbortController().signal,
        onDelta: () => undefined,
        onToolCall: (tc) => toolCalls.push(tc),
      },
    )

    expect(toolCalls.length).toBe(1)
    expect(toolCalls[0]).toMatchObject({
      name: 'apply_ops',
      input: {
        ops: [{ op: 'deleteBlocks', target: { blockIndexes: [2, 3] } }],
      },
    })
  })

  it('blocks a tool not listed in availableTools — allowlist enforcement (Blocker 4)', async () => {
    // Blocker 4: model must NOT escalate to ops not in availableTools
    const baseUrl = 'http://127.0.0.1:45679/v1'
    registerLocalVlmRuntime({
      status: vi.fn() as never,
      isRunnable: () => true,
      shouldFallbackFromAgy: () => true,
      ensureRunning: async () => ({ baseUrl, model: LOCAL_VLM_DEFAULT_MODEL }),
      testConnection: async () => ({ ok: true }),
      stop: async () => undefined,
    })

    vi.stubGlobal('fetch', (async (_input: RequestInfo | URL, init?: RequestInit) => {
      const parsed = JSON.parse(String(init?.body ?? '{}')) as Record<string, unknown>
      if (parsed.stream) {
        // Model tries to call apply_ops but availableTools only has read_blocks
        const sse = [
          `data: ${JSON.stringify({ choices: [{ delta: { content: '<tool_call>\n<function=setFont>\n<arguments>\n{"fontSize":12}\n</arguments>\n</function>\n</tool_call>' }, finish_reason: null }] })}\n\n`,
          `data: ${JSON.stringify({ choices: [{ delta: {}, finish_reason: 'stop' }] })}\n\n`,
          'data: [DONE]\n\n',
        ].join('')
        return new Response(sse, { status: 200, headers: { 'Content-Type': 'text/event-stream' } })
      }
      return new Response('{}', { status: 200 })
    }) as typeof fetch)

    const toolCalls: unknown[] = []
    await streamForProvider(
      'local-vlm',
      { apiKey: '', model: '' },
      'You are a document editing assistant',
      [{ role: 'user', text: 'just read the doc' }],
      // apply_ops is NOT in availableTools, so setFont (which maps to apply_ops) must be blocked
      [{ name: 'read_blocks', description: 'Read blocks', inputSchema: { type: 'object', properties: {} } }],
      512,
      {
        signal: new AbortController().signal,
        onDelta: () => undefined,
        onToolCall: (tc) => toolCalls.push(tc),
      },
    )

    expect(toolCalls.length).toBe(0)
  })
})

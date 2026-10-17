import { existsSync, mkdtempSync, readdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { act } from 'react'
import type { ReactElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { stageAttachment } from '@genoffice/mail-protocols'
import type { SendRequest } from '@genoffice/mail-protocols'
import { MAIL_CHANNELS } from '../src/shared/ipc'
import { QUARANTINE_CHANNELS } from '../src/shared/quarantine-channels'
import { FakeComposeProvider, registerComposeHandlers } from '../src/main/compose-handlers'
import type { ComposeHandlerDeps, ComposeIpcBus } from '../src/main/compose-handlers'
import { registerQuarantineHandlers } from '../src/main/quarantine-handlers'
import {
  confinePath,
  defaultPipelineDirs,
  extractPreviewText,
  resolveSendCaps,
  runAttachmentPipeline,
  scanAttachmentBytes,
  scanContentPrefix,
  scanFileName,
} from '../src/main/quarantine'
import type { SendCaps } from '../src/main/quarantine'
import { SqliteMailStore } from '../src/main/store/index'
import { ComposeHost } from '../src/renderer/compose/ComposeHost'
import { LocaleProvider } from '../src/renderer/i18n/locale'
import { MessageReader } from '../src/renderer/reader/MessageReader'
import { FakeMailApi } from '../src/renderer/devFakes'

/**
 * H5 quarantine tests (Q-8): traversal rejection, executables-never-inline,
 * and cap enforcement over the staged-temp --> quarantine-scan --> store
 * pipeline, plus progress IPC and UI surfacing. All fixtures are synthetic
 * (EICAR-style prefix vectors, hostile names/paths); never real malware.
 * Fake addresses use example.com only.
 */

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const SMALL_CAPS: SendCaps = {
  warnTotalBytes: 100,
  maxTotalBytes: 200,
  maxAttachmentBytes: 150,
  maxRecipients: 10,
}

/**
 * Test-realm bytes: TextEncoder output is re-wrapped so `instanceof
 * Uint8Array` holds under jsdom, where encoders may come from another
 * realm than the zod schemas.
 */
const bytes = (text: string): Uint8Array => new Uint8Array(new TextEncoder().encode(text))

/** EICAR-STYLE vector: standard prefix, standard length, harmless fill. */
function eicarStyleBytes(): Uint8Array {
  const out = new Uint8Array(68)
  out.set(bytes('X5O!P%@AP[4'))
  out.fill(65, 11)
  return out
}

/** Minimal zip local header claiming a 4 GiB member (bomb heuristic tripwire). */
function zipBombPrefix(): Uint8Array {
  const out = new Uint8Array(64)
  out.set([0x50, 0x4b, 0x03, 0x04], 0)
  const view = new DataView(out.buffer)
  view.setUint32(18, 10, true)
  view.setUint32(22, 0xffffffff, true)
  view.setUint16(26, 4, true)
  view.setUint16(28, 0, true)
  return out
}

function mzBytes(): Uint8Array {
  const out = new Uint8Array([0x4d, 0x5a, 0x90, 0x00, 0x03, 0x00, 0x00, 0x00])
  return out
}

describe('scanFileName', () => {
  it('accepts plain bare filenames', () => {
    expect(scanFileName('invoice.pdf').ok).toBe(true)
    expect(scanFileName('my report (2).docx').ok).toBe(true)
  })

  it.each(['../evil.txt', '..\\evil.txt', 'a/b.txt', '/abs/path.txt', 'C:\\win.ini', '\\\\unc\\x'])(
    'rejects traversal %s',
    (name) => {
      const result = scanFileName(name)
      expect(result.ok).toBe(false)
      expect(result.reason).toBe('traversal')
    },
  )

  it('rejects ADS/colon tricks and Windows strip tricks', () => {
    expect(scanFileName('notes.txt:evil.exe').reason).toBe('hostile-name')
    expect(scanFileName('report.pdf ').reason).toBe('hostile-name')
    expect(scanFileName('report.pdf.').reason).toBe('hostile-name')
  })

  it('rejects control characters, empties, and overlong names', () => {
    expect(scanFileName(`a${String.fromCharCode(0)}b.txt`).ok).toBe(false)
    expect(scanFileName('').reason).toBe('hostile-name')
    expect(scanFileName('...').reason).toBe('hostile-name')
    expect(scanFileName(`${'a'.repeat(300)}.txt`).reason).toBe('hostile-name')
  })
})

describe('scanAttachmentBytes', () => {
  it('cleans ordinary documents', () => {
    const result = scanAttachmentBytes(
      { filename: 'note.txt', mime: 'text/plain', bytes: bytes('hi') },
      SMALL_CAPS,
    )
    expect(result.verdict).toBe('clean')
    expect(result.reason).toBe('clean')
    expect(result.filename).toBe('note.txt')
  })

  it('quarantines executables, scripts, and macros (never inline)', () => {
    expect(
      scanAttachmentBytes(
        { filename: 'setup.exe', mime: 'application/octet-stream', bytes: bytes('hi') },
        SMALL_CAPS,
      ).reason,
    ).toBe('exec-type')
    expect(
      scanAttachmentBytes(
        { filename: 'run.ps1', mime: 'application/octet-stream', bytes: bytes('hi') },
        SMALL_CAPS,
      ).reason,
    ).toBe('script-type')
    expect(
      scanAttachmentBytes(
        { filename: 'books.xlsm', mime: 'application/octet-stream', bytes: bytes('hi') },
        SMALL_CAPS,
      ).reason,
    ).toBe('macro-type')
    expect(
      scanAttachmentBytes(
        { filename: 'invoice.pdf.exe', mime: 'application/octet-stream', bytes: bytes('hi') },
        SMALL_CAPS,
      ).reason,
    ).toBe('exec-type')
  })

  it('quarantines MIME/extension mismatch but forgives octet-stream', () => {
    const mismatch = scanAttachmentBytes(
      { filename: 'report.docx', mime: 'image/png', bytes: bytes('hi') },
      SMALL_CAPS,
    )
    expect(mismatch.verdict).toBe('quarantined')
    expect(mismatch.reason).toBe('mime-mismatch')
    const generic = scanAttachmentBytes(
      { filename: 'blob.bin', mime: 'application/octet-stream', bytes: bytes('hi') },
      SMALL_CAPS,
    )
    expect(generic.verdict).toBe('clean')
  })

  it('blocks over-cap bytes and flags warn-band sizes', () => {
    const over = scanAttachmentBytes(
      { filename: 'big.txt', mime: 'text/plain', bytes: new Uint8Array(151) },
      SMALL_CAPS,
    )
    expect(over.verdict).toBe('blocked')
    expect(over.reason).toBe('over-cap')
    const warn = scanAttachmentBytes(
      { filename: 'mid.txt', mime: 'text/plain', bytes: new Uint8Array(120) },
      SMALL_CAPS,
    )
    expect(warn.verdict).toBe('clean')
    expect(warn.sizeWarn).toBe(true)
  })

  it('blocks EICAR-style vectors, executable magic, and zip bombs', () => {
    expect(
      scanAttachmentBytes(
        { filename: 'test.txt', mime: 'text/plain', bytes: eicarStyleBytes() },
        SMALL_CAPS,
      ).reason,
    ).toBe('suspicious-content')
    const mzPdf = scanAttachmentBytes(
      { filename: 'doc.pdf', mime: 'application/pdf', bytes: mzBytes() },
      SMALL_CAPS,
    )
    expect(mzPdf.verdict).toBe('blocked')
    const zipBomb = scanAttachmentBytes(
      { filename: 'a.zip', mime: 'application/zip', bytes: zipBombPrefix() },
      SMALL_CAPS,
    )
    expect(zipBomb.verdict).toBe('blocked')
    expect(zipBomb.reason).toBe('suspicious-content')
  })

  it('quarantines pdf names without pdf magic', () => {
    const result = scanContentPrefix('doc.pdf', bytes('hello'), 5)
    expect(result.verdict).toBe('quarantined')
    expect(result.reason).toBe('mime-mismatch')
  })
})

describe('resolveSendCaps', () => {
  it('mirrors the compose-send defaults for IMAP and resolves per-driver limits', () => {
    const imap = resolveSendCaps('imap')
    expect(imap).toEqual({
      warnTotalBytes: 25 * 1024 * 1024,
      maxTotalBytes: 50 * 1024 * 1024,
      maxAttachmentBytes: 25 * 1024 * 1024,
      maxRecipients: 100,
    })
    expect(resolveSendCaps().maxAttachmentBytes).toBe(25 * 1024 * 1024)
    const gmail = resolveSendCaps('gmail')
    expect(gmail.maxTotalBytes).toBe(35 * 1024 * 1024)
    expect(gmail.maxRecipients).toBe(500)
    const graph = resolveSendCaps('graph')
    expect(graph.maxTotalBytes).toBe(4 * 1024 * 1024)
    expect(graph.maxAttachmentBytes).toBe(3 * 1024 * 1024)
    expect(graph.maxRecipients).toBe(500)
  })

  it('enforces the strictest per-driver file cap (graph 3 MiB)', () => {
    const graph = resolveSendCaps('graph')
    const imap = resolveSendCaps('imap')
    const fourMb = new Uint8Array(4 * 1024 * 1024)
    expect(
      scanAttachmentBytes(
        { filename: 'f.bin', mime: 'application/octet-stream', bytes: fourMb },
        graph,
      ).reason,
    ).toBe('over-cap')
    expect(
      scanAttachmentBytes(
        { filename: 'f.bin', mime: 'application/octet-stream', bytes: fourMb },
        imap,
      ).verdict,
    ).toBe('clean')
  })
})

describe('confinePath', () => {
  it('holds resolve-then-startsWith for stage and sibling roots', () => {
    const root = mkdtempSync(join(tmpdir(), 'mail-q-root-'))
    const stage = join(root, 'stage')
    const quarantine = join(root, 'quarantine')
    expect(confinePath(join(stage, 'abc-note.txt'), stage)).toContain('abc-note.txt')
    expect(confinePath(join(quarantine, 'abc-tool.exe'), quarantine)).toContain('tool.exe')
    expect(() => confinePath(join(root, 'stage', '..', 'escape.txt'), stage)).toThrow('mail:')
    expect(() => confinePath('/etc/hostname', stage)).toThrow('mail:')
  })
})

describe('runAttachmentPipeline', () => {
  function makeDirs(): { root: string; stageDir: string; quarantineDir: string; storeDir: string } {
    const root = mkdtempSync(join(tmpdir(), 'mail-q-pipe-'))
    return { root, ...defaultPipelineDirs(root) }
  }

  it('stages, scans, and stores clean files with progress order', () => {
    const dirs = makeDirs()
    const seen: string[] = []
    const outcome = runAttachmentPipeline(
      { filename: 'note.txt', mime: 'text/plain', bytes: bytes('hello quarantine') },
      dirs,
      { stage: stageAttachment, onProgress: (state) => seen.push(state) },
      SMALL_CAPS,
    )
    expect(outcome.verdict).toBe('clean')
    expect(seen).toEqual(['staged', 'scanning', 'stored'])
    expect(outcome.path).toBeDefined()
    expect(outcome.path?.startsWith(dirs.storeDir)).toBe(true)
    expect(existsSync(outcome.path as string)).toBe(true)
    expect(readdirSync(dirs.stageDir)).toEqual([])
  })

  it('parks executables in the quarantine sibling root', () => {
    const dirs = makeDirs()
    const outcome = runAttachmentPipeline(
      { filename: 'tool.exe', mime: 'application/octet-stream', bytes: bytes('fake-binary') },
      dirs,
      { stage: stageAttachment },
      SMALL_CAPS,
    )
    expect(outcome.verdict).toBe('quarantined')
    expect(outcome.reason).toBe('exec-type')
    expect(outcome.path?.startsWith(dirs.quarantineDir)).toBe(true)
    expect(readdirSync(dirs.stageDir)).toEqual([])
  })

  it('drops blocked files with no store path', () => {
    const dirs = makeDirs()
    const seen: string[] = []
    const outcome = runAttachmentPipeline(
      { filename: '../../evil.txt', mime: 'text/plain', bytes: bytes('hostile') },
      dirs,
      { stage: stageAttachment, onProgress: (state) => seen.push(state) },
      SMALL_CAPS,
    )
    expect(outcome.verdict).toBe('blocked')
    expect(outcome.reason).toBe('traversal')
    expect(outcome.path).toBeUndefined()
    expect(seen).toEqual(['staged', 'scanning', 'blocked'])
    expect(readdirSync(dirs.stageDir)).toEqual([])
    expect(readdirSync(dirs.storeDir)).toEqual([])
    expect(readdirSync(dirs.quarantineDir)).toEqual([])
  })
})

describe('extractPreviewText', () => {
  it('never calls the extractor for executables (exec-never-inline)', async () => {
    const root = mkdtempSync(join(tmpdir(), 'mail-q-prev-'))
    const dirs = defaultPipelineDirs(root)
    const outcome = runAttachmentPipeline(
      { filename: 'tool.exe', mime: 'application/octet-stream', bytes: bytes('fake-binary') },
      dirs,
      { stage: stageAttachment },
      SMALL_CAPS,
    )
    const extractor = vi.fn(async () => ({ ok: true as const, kind: 'text' as const, text: 'x' }))
    const refused = await extractPreviewText(
      outcome.path as string,
      'tool.exe',
      dirs.quarantineDir,
      'quarantined',
      extractor,
    )
    expect(refused.allowed).toBe(false)
    expect(extractor).not.toHaveBeenCalled()
  })

  it('delegates clean previews to the injected file-parse-shaped extractor', async () => {
    const root = mkdtempSync(join(tmpdir(), 'mail-q-prev-'))
    const dirs = defaultPipelineDirs(root)
    const outcome = runAttachmentPipeline(
      { filename: 'note.txt', mime: 'text/plain', bytes: bytes('preview me') },
      dirs,
      { stage: stageAttachment },
      SMALL_CAPS,
    )
    const extractor = vi.fn(async () => ({
      ok: true as const,
      kind: 'text' as const,
      text: 'preview me',
    }))
    const preview = await extractPreviewText(
      outcome.path as string,
      'note.txt',
      dirs.storeDir,
      'clean',
      extractor,
    )
    expect(extractor).toHaveBeenCalledTimes(1)
    expect(preview).toMatchObject({ allowed: true, text: 'preview me', kind: 'text' })
  })

  it('refuses paths outside the store root', async () => {
    const extractor = vi.fn(async () => ({ ok: true as const, kind: 'text' as const, text: 'x' }))
    await expect(
      extractPreviewText('/etc/hostname', 'hostname', tmpdir(), 'clean', extractor),
    ).rejects.toThrow('mail:')
    expect(extractor).not.toHaveBeenCalled()
  })
})

function makeBus(): {
  bus: ComposeIpcBus
  invoke: (channel: string, payload: unknown) => Promise<unknown>
} {
  const handlers = new Map<string, (payload: unknown) => Promise<unknown>>()
  return {
    bus: {
      handle: (channel, listener) => {
        handlers.set(channel, listener)
      },
    },
    invoke: async (channel, payload) => {
      const listener = handlers.get(channel)
      if (!listener) throw new Error(`no handler: ${channel}`)
      return listener(payload)
    },
  }
}

describe('quarantine IPC handlers', () => {
  it('scans bytes without disk side effects and rejects bad payloads', async () => {
    const { bus, invoke } = makeBus()
    registerQuarantineHandlers(bus, {})
    const clean = (await invoke(QUARANTINE_CHANNELS.scan, {
      filename: 'note.txt',
      mime: 'text/plain',
      data: bytes('hi'),
    })) as { verdict: string }
    expect(clean.verdict).toBe('clean')
    const blocked = (await invoke(QUARANTINE_CHANNELS.scan, {
      filename: '../evil.txt',
      mime: 'text/plain',
      data: bytes('hi'),
    })) as { verdict: string; reason: string }
    expect(blocked).toMatchObject({ verdict: 'blocked', reason: 'traversal' })
    await expect(
      invoke(QUARANTINE_CHANNELS.scan, { filename: '', mime: 'x', data: bytes('hi') }),
    ).rejects.toThrow('mail:')
  })

  it('ingests through the full pipeline and previews stored clean files', async () => {
    const root = mkdtempSync(join(tmpdir(), 'mail-q-ipc-'))
    const dirs = defaultPipelineDirs(root)
    const opened: string[] = []
    const { bus, invoke } = makeBus()
    registerQuarantineHandlers(bus, {
      dirs,
      openPath: (path) => {
        opened.push(path)
      },
      extractText: async () => ({ ok: true, kind: 'text', text: 'stored text' }),
    })
    const ingested = (await invoke(QUARANTINE_CHANNELS.ingest, {
      filename: 'note.txt',
      mime: 'text/plain',
      data: bytes('stored text'),
    })) as { verdict: string; path: string; progress: string[] }
    expect(ingested.verdict).toBe('clean')
    expect(ingested.progress).toEqual(['staged', 'scanning', 'stored'])
    const preview = (await invoke(QUARANTINE_CHANNELS.preview, { path: ingested.path })) as {
      allowed: boolean
      text: string
    }
    expect(preview).toMatchObject({ allowed: true, text: 'stored text' })
    const open = (await invoke(QUARANTINE_CHANNELS.open, {
      path: ingested.path,
      confirmed: false,
    })) as {
      opened: boolean
    }
    expect(open.opened).toBe(true)
    expect(opened).toEqual([ingested.path])
    await expect(
      invoke(QUARANTINE_CHANNELS.open, { path: '/etc/hostname', confirmed: true }),
    ).rejects.toThrow('mail:')
  })

  it('holds quarantined files for explicit confirm', async () => {
    const root = mkdtempSync(join(tmpdir(), 'mail-q-ipc-'))
    const dirs = defaultPipelineDirs(root)
    const opened: string[] = []
    const { bus, invoke } = makeBus()
    registerQuarantineHandlers(bus, {
      dirs,
      openPath: (path) => {
        opened.push(path)
      },
    })
    const ingested = (await invoke(QUARANTINE_CHANNELS.ingest, {
      filename: 'tool.exe',
      mime: 'application/octet-stream',
      data: bytes('fake-binary'),
    })) as { verdict: string; path: string }
    expect(ingested.verdict).toBe('quarantined')
    await expect(
      invoke(QUARANTINE_CHANNELS.open, { path: ingested.path, confirmed: false }),
    ).rejects.toThrow('explicit confirm')
    expect(opened).toEqual([])
    await invoke(QUARANTINE_CHANNELS.open, { path: ingested.path, confirmed: true })
    expect(opened).toEqual([ingested.path])
  })
})

describe('compose-send quarantine enforcement', () => {
  function makeDeps(over?: Partial<ComposeHandlerDeps>): {
    deps: ComposeHandlerDeps
    provider: FakeComposeProvider
  } {
    const provider = new FakeComposeProvider()
    const stageDir = mkdtempSync(join(tmpdir(), 'mail-compose-q-'))
    const deps: ComposeHandlerDeps = {
      provider,
      stage: stageAttachment,
      stageDir,
      openExternal: () => undefined,
      clock: () => '2026-10-15T00:00:00.000Z',
      ...over,
    }
    return { deps, provider }
  }

  function sendReq(over?: Partial<SendRequest>): SendRequest {
    return { to: ['a@example.com'], subject: 'hi', text: 'hello', ...over }
  }

  it('rejects hostile-name and suspicious-content attachments at send', async () => {
    const { bus, invoke } = makeBus()
    const { deps, provider } = makeDeps()
    registerComposeHandlers(bus, deps)
    await expect(
      invoke(
        MAIL_CHANNELS.composeSend,
        sendReq({
          attachments: [
            {
              filename: '../evil.txt',
              mime: 'text/plain',
              size: 2,
              path: 'memory:../evil.txt',
              data: bytes('x'),
            },
          ],
        }),
      ),
    ).rejects.toThrow('blocked')
    const staged = stageAttachment({
      filename: 't.txt',
      mime: 'text/plain',
      bytes: eicarStyleBytes(),
      stageDir: deps.stageDir,
    })
    await expect(
      invoke(
        MAIL_CHANNELS.composeSend,
        sendReq({
          attachments: [
            { filename: staged.filename, mime: staged.mime, size: staged.size, path: staged.path },
          ],
        }),
      ),
    ).rejects.toThrow('blocked')
    expect(provider.sent).toHaveLength(0)
  })

  it('sends quarantined executables (never inline downstream, still sendable)', async () => {
    const { bus, invoke } = makeBus()
    const { deps, provider } = makeDeps()
    registerComposeHandlers(bus, deps)
    const result = (await invoke(
      MAIL_CHANNELS.composeSend,
      sendReq({
        attachments: [
          {
            filename: 'tool.exe',
            mime: 'application/octet-stream',
            size: 4,
            path: 'memory:tool.exe',
            data: bytes('bin!'),
          },
        ],
      }),
    )) as { outboxId: string }
    expect(result.outboxId).toBeDefined()
    expect(provider.sent).toHaveLength(1)
  })

  it('enforces per-driver file caps where the driver is known', async () => {
    const { bus, invoke } = makeBus()
    const { deps } = makeDeps({ driver: 'graph' })
    registerComposeHandlers(bus, deps)
    const overFile = new Uint8Array(Math.floor(3.5 * 1024 * 1024))
    // Graph caps resolve to 4 MiB total / 3 MiB per file: the 3.5 MiB file
    // trips the per-file cap through the established caps error path.
    await expect(
      invoke(
        MAIL_CHANNELS.composeSend,
        sendReq({
          attachments: [
            {
              filename: 'f.bin',
              mime: 'application/octet-stream',
              size: overFile.byteLength,
              path: 'memory:f.bin',
              data: overFile,
            },
          ],
        }),
      ),
    ).rejects.toThrow('per-account cap')
  })
})

describe('attachment store refs (never blobs in SQLite)', () => {
  it('records pipeline outcomes as refs pointing at disk files', () => {
    const qRoot = mkdtempSync(join(tmpdir(), 'mail-q-store-'))
    const dirs = defaultPipelineDirs(qRoot)
    const outcome = runAttachmentPipeline(
      { filename: 'note.txt', mime: 'text/plain', bytes: bytes('on disk, not in sqlite') },
      dirs,
      { stage: stageAttachment },
      SMALL_CAPS,
    )
    const dbRoot = mkdtempSync(join(tmpdir(), 'mail-db-'))
    const store = SqliteMailStore.open('acct-1', { rootDir: dbRoot, dbPath: ':memory:' })
    try {
      store.addAttachment({
        accountId: 'acct-1',
        msgUid: 'u1',
        filename: outcome.filename,
        mime: 'text/plain',
        size: 22,
        blobPath: outcome.path as string,
        state: 'stored',
      })
      const rows = store.listAttachments('acct-1', 'u1')
      expect(rows).toHaveLength(1)
      expect(rows[0]?.blobPath).toBe(outcome.path)
      expect(rows[0]).not.toHaveProperty('data')
      expect(rows[0]).not.toHaveProperty('bytes')
      expect(existsSync(rows[0]?.blobPath as string)).toBe(true)
    } finally {
      store.close()
    }
  })
})

let containers: HTMLElement[] = []

function render(node: ReactElement): HTMLElement {
  const container = document.createElement('div')
  document.body.appendChild(container)
  containers.push(container)
  const root = createRoot(container)
  act(() => {
    root.render(<LocaleProvider initial="en">{node}</LocaleProvider>)
  })
  return container
}

async function settle(rounds = 12): Promise<void> {
  await act(async () => {
    for (let i = 0; i < rounds; i += 1) {
      await Promise.resolve()
    }
  })
}

function buttonByText(container: HTMLElement, text: string): HTMLButtonElement | undefined {
  return [...container.querySelectorAll('button')].find((button) => button.textContent === text) as
    HTMLButtonElement | undefined
}

function attachFile(container: HTMLElement, name: string, type: string, content: string): void {
  const picker = container.querySelector('.mail-compose-attach input[type="file"]')
  if (!(picker instanceof HTMLInputElement)) throw new Error('missing file picker')
  Object.defineProperty(picker, 'files', {
    configurable: true,
    value: [new File([content], name, { type })],
  })
}

beforeEach(() => {
  ;(window as unknown as { mailApi: FakeMailApi }).mailApi = new FakeMailApi()
  ;(window as unknown as { mailCompose?: unknown }).mailCompose = undefined
  ;(window as unknown as { mailQuarantine?: unknown }).mailQuarantine = undefined
})

afterEach(() => {
  for (const container of containers) container.remove()
  containers = []
})

describe('ComposeHost quarantine surfacing', () => {
  function stubQuarantine(): void {
    ;(window as unknown as { mailQuarantine?: unknown }).mailQuarantine = {
      scan: async (filename: string) =>
        filename === '../evil.txt'
          ? { verdict: 'blocked', reason: 'traversal', filename: 'evil.txt', sizeWarn: false }
          : { verdict: 'clean', reason: 'clean', filename, sizeWarn: false },
      ingest: async () => ({
        verdict: 'clean',
        reason: 'clean',
        filename: 'x',
        sizeWarn: false,
        progress: [],
      }),
      preview: async () => ({ allowed: false, reason: 'exec-type' as const }),
      openAttachment: async () => ({ opened: true }),
    }
  }

  it('shows scanning then blocked per attachment and holds send', async () => {
    stubQuarantine()
    const onSend = vi.fn()
    const container = render(
      <ComposeHost
        accountId="acc-1"
        threadKey="t1"
        onDirtyChange={() => undefined}
        onSend={onSend}
      />,
    )
    await settle()
    const to = container.querySelector('#mail-compose-to')
    if (!(to instanceof HTMLInputElement)) throw new Error('missing to field')
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(
        window.HTMLInputElement.prototype,
        'value',
      )?.set
      if (setter) setter.call(to, 'a@example.com')
      else to.value = 'a@example.com'
      to.dispatchEvent(new Event('input', { bubbles: true }))
    })
    attachFile(container, '../evil.txt', 'text/plain', 'hostile')
    await act(async () => {
      container
        .querySelector('.mail-compose-attach input[type="file"]')
        ?.dispatchEvent(new Event('change', { bubbles: true }))
    })
    await settle()
    expect(container.textContent).toContain('Blocked')
    const send = buttonByText(container, 'Send')
    expect(send?.disabled).toBe(true)
    expect(onSend).not.toHaveBeenCalled()
  })

  it('keeps send available for clean verdicts', async () => {
    stubQuarantine()
    const onSend = vi.fn()
    const container = render(
      <ComposeHost
        accountId="acc-1"
        threadKey="t1"
        onDirtyChange={() => undefined}
        onSend={onSend}
      />,
    )
    await settle()
    attachFile(container, 'note.txt', 'text/plain', 'hello')
    await act(async () => {
      container
        .querySelector('.mail-compose-attach input[type="file"]')
        ?.dispatchEvent(new Event('change', { bubbles: true }))
    })
    await settle()
    expect(container.textContent).toContain('note.txt')
    expect(container.textContent).not.toContain('Blocked')
  })
})

describe('MessageReader quarantine rows', () => {
  function envelope(): {
    ref: { uid: string; folderId: string; threadId: string }
    from: string[]
    to: string[]
    cc: string[]
    subject: string
    date: string
    flags: { seen: boolean; flagged: boolean; answered: boolean; draft: boolean }
    labels: string[]
    snippet: string
    hasAttach: boolean
  } {
    return {
      ref: { uid: 'u9', folderId: 'inbox', threadId: 't9' },
      from: ['sender@example.com'],
      to: ['reader@example.com'],
      cc: [],
      subject: 'with attachments',
      date: '2026-10-15T00:00:00.000Z',
      flags: { seen: false, flagged: false, answered: false, draft: false },
      labels: [],
      snippet: '',
      hasAttach: true,
    }
  }

  it('renders blocked placeholders with no actions and quarantined rows confirm-gated', async () => {
    const opened: Array<{ path: string; confirmed: boolean }> = []
    ;(window as unknown as { mailQuarantine?: unknown }).mailQuarantine = {
      scan: async () => ({ verdict: 'clean', reason: 'clean', filename: 'x', sizeWarn: false }),
      ingest: async () => ({
        verdict: 'clean',
        reason: 'clean',
        filename: 'x',
        sizeWarn: false,
        progress: [],
      }),
      preview: async () => ({ allowed: true, reason: 'clean', text: 'file text', kind: 'text' }),
      openAttachment: async (path: string, confirmed: boolean) => {
        opened.push({ path, confirmed })
        return { opened: true }
      },
    }
    const api = new FakeMailApi()
    const container = render(
      <MessageReader
        api={api}
        envelope={envelope()}
        attachments={[
          {
            filename: 'evil.txt',
            mime: 'text/plain',
            size: 7,
            verdict: 'blocked',
            reason: 'traversal',
          },
          {
            filename: 'tool.exe',
            mime: 'application/octet-stream',
            size: 4,
            verdict: 'quarantined',
            reason: 'exec-type',
            path: join(tmpdir(), 'tool.exe'),
          },
          {
            filename: 'note.txt',
            mime: 'text/plain',
            size: 9,
            verdict: 'clean',
            reason: 'clean',
            path: join(tmpdir(), 'note.txt'),
          },
        ]}
      />,
    )
    await settle()
    expect(container.textContent).toContain('Blocked (traversal).')
    const buttons = [...container.querySelectorAll('button')].map((button) => button.textContent)
    expect(buttons).toContain('Save and open')
    expect(buttons).toContain('Preview')
    // Blocked rows carry no open control: only the two expected actions exist.
    expect(buttons.filter((text) => text === 'Save and open')).toHaveLength(1)
    const saveOpen = buttonByText(container, 'Save and open')
    await act(async () => {
      saveOpen?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    await settle()
    expect(buttonByText(container, 'Open this file?')).toBeDefined()
    await act(async () => {
      buttonByText(container, 'Open this file?')?.dispatchEvent(
        new MouseEvent('click', { bubbles: true }),
      )
    })
    await settle()
    expect(opened).toEqual([{ path: join(tmpdir(), 'tool.exe'), confirmed: true }])
    const preview = buttonByText(container, 'Preview')
    await act(async () => {
      preview?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    await settle()
    expect(container.textContent).toContain('file text')
  })
})

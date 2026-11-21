import { describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({
  session: { defaultSession: {} },
  app: { getAppPath: () => '/tmp/nowhere/apps/shell', isPackaged: false },
  // capture.ts self-installs over this noop; docs-main's module-scope
  // registrations (pulled in via web/index → bridge) land in its tables
  ipcMain: { handle: () => {}, on: () => {}, once: () => {}, removeHandler: () => {} },
}))
vi.mock('../../docs/src/main/docs-main', () => ({
  queueDocsPendingOpen: () => {},
  markDocsNewBlank: () => {},
}))
vi.mock('../../sheets/src/main/sheets-main', () => ({
  createSheetsWebTab: () => {},
  setSheetsNewBlank: () => {},
}))
vi.mock('../../slides/src/main/slides-main', () => ({ trackWebSlidesTab: () => {} }))
vi.mock('../../markdown/src/main/markdown-main', () => ({ trackWebTab: () => {} }))
vi.mock('../../pdf/src/main/pdf-main', () => ({ trackWebTab: () => {} }))
vi.mock('../../html/src/main/html-main', () => ({ trackWebTab: () => {} }))
vi.mock('../../docs/src/main/lazy-media', () => ({ readLazyMedia: async () => null }))
vi.mock('../../docs/src/main/byte-handoff', () => ({ takeHandoff: () => null }))
vi.mock('../../markdown/src/main/markdown-main', () => ({
  serveMarkdownAsset: async () => new Response(null),
}))
vi.mock('../../html/src/main/html-main', () => ({
  serveHtmlAsset: async () => new Response(null),
  previewEntryFor: () => null,
}))
vi.mock('../../html/src/main/preview-document', () => ({
  buildPreviewDocument: (t: string) => t,
}))

import { originHostOk, tokenValid } from '../src/web/static'
import { parseWebServerArgv } from '../src/web/index'
import { WebTabContents } from '../src/web/web-contents'

const BOUND = { host: '127.0.0.1', port: 8931 }

function reqWith(
  headers: Record<string, string>,
  url = '/docs/',
): import('node:http').IncomingMessage {
  return { headers, url } as unknown as import('node:http').IncomingMessage
}

describe('web request guard', () => {
  it('allows requests without an Origin (curl, top-level navigations)', () => {
    expect(originHostOk(reqWith({ host: '127.0.0.1:8931' }), BOUND)).toBe(true)
  })

  it('requires a plausible Host header', () => {
    expect(originHostOk(reqWith({}), BOUND)).toBe(false)
    expect(originHostOk(reqWith({ host: 'bad\0host' }), BOUND)).toBe(false)
  })

  it('accepts an Origin that matches a loopback or the bound host', () => {
    expect(
      originHostOk(reqWith({ host: '127.0.0.1:8931', origin: 'http://127.0.0.1:8931' }), BOUND),
    ).toBe(true)
    expect(originHostOk(reqWith({ host: 'localhost:8931' }), BOUND)).toBe(true)
    expect(
      originHostOk(reqWith({ host: 'myhost.lan:8931', origin: 'http://myhost.lan:8931' }), {
        host: 'myhost.lan',
        port: 8931,
      }),
    ).toBe(true)
  })

  it('refuses a rebound or foreign Host even when Origin agrees', () => {
    // the DNS-rebinding case: evil.com resolves at 127.0.0.1, so Origin and
    // Host agree — but the hostname is not one this server is served on
    expect(
      originHostOk(reqWith({ host: 'evil.com:8931', origin: 'http://evil.com:8931' }), BOUND),
    ).toBe(false)
    expect(originHostOk(reqWith({ host: 'evil.com:8931' }), BOUND)).toBe(false)
  })

  it('refuses cross-site origins (another page driving the local service)', () => {
    expect(
      originHostOk(reqWith({ host: '127.0.0.1:8931', origin: 'https://evil.example' }), BOUND),
    ).toBe(false)
    expect(
      originHostOk(reqWith({ host: '127.0.0.1:8931', origin: 'http://127.0.0.1:9999' }), BOUND),
    ).toBe(false)
  })

  it('refuses a Host with the wrong or missing port', () => {
    expect(originHostOk(reqWith({ host: '127.0.0.1:9999' }), BOUND)).toBe(false)
    expect(originHostOk(reqWith({ host: '127.0.0.1' }), BOUND)).toBe(false)
  })

  it('refuses malformed or non-http origins', () => {
    expect(originHostOk(reqWith({ host: '127.0.0.1:8931', origin: 'not a url' }), BOUND)).toBe(
      false,
    )
    expect(originHostOk(reqWith({ host: '127.0.0.1:8931', origin: 'file:///etc' }), BOUND)).toBe(
      false,
    )
  })

  it('keeps requiring the port-scoped token for cookies and query', () => {
    const req = reqWith({ host: 'h', cookie: 'genoffice_web_token_8931=t' })
    expect(tokenValid(req, 't', 8931)).toBe(true)
    expect(tokenValid(reqWith({ host: 'h' }, '/docs/?token=t'), 't', 8931)).toBe(true)
    expect(
      tokenValid(reqWith({ host: 'h', cookie: 'genoffice_web_token_8931=x' }), 't', 8931),
    ).toBe(false)
    // a cookie scoped to another port does not authenticate here
    expect(
      tokenValid(reqWith({ host: 'h', cookie: 'genoffice_web_token_9999=t' }), 't', 8931),
    ).toBe(false)
  })
})

describe('web-server argv', () => {
  it('a bare --web-token generates instead of eating the next flag', () => {
    const parsed = parseWebServerArgv(['app', '--web-server', '--web-token', '--web-port', '9000'])
    expect(parsed.enabled).toBe(true)
    expect(parsed.overrides.token).toMatch(/^[0-9a-f]{48}$/)
    expect(parsed.overrides.port).toBe(9000)
  })

  it('--web-token with a value pins it', () => {
    const parsed = parseWebServerArgv(['--web-server', '--web-token', 'sekret'])
    expect(parsed.overrides.token).toBe('sekret')
  })

  it('a token-less invocation stays token-hungry at startup', () => {
    const parsed = parseWebServerArgv(['--web-server'])
    expect(parsed.enabled).toBe(true)
    expect(parsed.overrides.token).toBeUndefined()
  })
})

describe('web tab once()', () => {
  it('fires once and removes itself from the listener set', () => {
    const tab = new WebTabContents('/docs/?path=/x.docx')
    const calls: string[] = []
    tab.once('destroyed', () => calls.push('a'))
    tab.on('destroyed', () => calls.push('keep'))
    tab.destroy()
    tab.destroy()
    expect(calls).toEqual(['a', 'keep'])
  })

  it('an once registered after destroy fires immediately once', () => {
    const tab = new WebTabContents('/docs/')
    tab.destroy()
    const calls: string[] = []
    tab.once('destroyed', () => calls.push('late'))
    expect(calls).toEqual(['late'])
  })
})

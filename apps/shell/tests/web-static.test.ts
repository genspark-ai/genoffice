import { afterEach, describe, expect, it, vi } from 'vitest'
import { mkdtempSync, mkdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

vi.mock('electron', () => ({
  app: { getAppPath: () => '/tmp/nowhere/apps/shell', isPackaged: false },
}))
vi.mock('../../docs/src/main/lazy-media', () => ({
  readLazyMedia: async (url: string) =>
    url === 'genoffice-docx-media://abc/media/image1.png'
      ? { body: new Uint8Array([1, 2, 3]), mime: 'image/png' }
      : null,
}))
vi.mock('../../markdown/src/main/markdown-main', () => ({
  serveMarkdownAsset: async (url: string) =>
    new Response(`asset:${url}`, { status: url.includes('ok.png') ? 200 : 403 }),
}))
vi.mock('../../html/src/main/html-main', () => ({
  serveHtmlAsset: async (url: string) => new Response(`htmlasset:${url}`),
  previewEntryFor: (wcId: number) =>
    wcId === 42 ? { text: '<p>preview</p>', baseHref: null } : null,
}))
vi.mock('../../html/src/main/preview-document', () => ({
  buildPreviewDocument: (text: string) => `<html>${text}</html>`,
}))

import { serveWebStatic, webRendererRoots, type RendererRoots } from '../src/web/static'

type Res = { status: number; body: string | Buffer; headers: Record<string, string | undefined> }

async function get(
  roots: RendererRoots,
  url: string,
  token: string | null = null,
  cookie = '',
): Promise<Res> {
  const req = { url, headers: { cookie } } as never
  const resMock = {
    headers: {} as Record<string, string | undefined>,
    head: null as unknown as [number, Record<string, string>] | null,
    chunks: [] as Buffer[],
    writeHead(status: number, headers?: Record<string, string>) {
      this.head = [status, headers ?? {}]
      for (const [k, v] of Object.entries(headers ?? {})) this.headers[k.toLowerCase()] = v
    },
    setHeader(k: string, v: string) {
      this.headers[k.toLowerCase()] = v
    },
    end(body?: unknown) {
      if (typeof body === 'string' || Buffer.isBuffer(body)) this.chunks.push(Buffer.from(body))
    },
  }
  await serveWebStatic(req, resMock as never, { roots, token })
  const res: Res = {
    status: resMock.head?.[0] ?? 0,
    body: Buffer.concat(resMock.chunks),
    headers: resMock.headers,
  }
  return res
}

describe('web static serving', () => {
  let dir: string
  afterEach(() => {
    dir = ''
  })

  function fixtureRoot(): RendererRoots {
    dir = mkdtempSync(join(tmpdir(), 'webstatic-'))
    writeFileSync(
      join(dir, 'index.html'),
      '<!doctype html><html><head><meta http-equiv="Content-Security-Policy" content="default-src \'self\'; connect-src \'self\'">' +
        '</head><body><script type="module" src="./assets/index-abc.js"></script></body></html>',
    )
    mkdirSync(join(dir, 'assets'))
    writeFileSync(
      join(dir, 'assets', 'index-abc.js'),
      'const s = "genoffice-docx-media://h/p"; s.startsWith("genoffice-docx-media://")',
    )
    return { docs: dir }
  }

  it('serves the launcher at /', async () => {
    const res = await get({}, '/')
    expect(res.status).toBe(200)
    expect(String(res.body)).toContain('GenOffice Web')
  })

  it('serves index.html with the bridge injected and CSP widened', async () => {
    const roots = fixtureRoot()
    const res = await get(roots, '/docs/')
    expect(res.status).toBe(200)
    const html = String(res.body)
    expect(html).toContain('<script src="/__web/client.js"></script>')
    expect(html).toContain('preload-docs.js')
    expect(html.indexOf('/__web/client.js')).toBeLessThan(html.indexOf('type="module"'))
    expect(html).toMatch(/connect-src [^"]*ws:/)
  })

  it('rewrites the docs media scheme to a same-origin path', async () => {
    const roots = fixtureRoot()
    const res = await get(roots, '/docs/assets/index-abc.js')
    const js = String(res.body)
    expect(js).not.toContain('genoffice-docx-media')
    expect(js).toContain('/__media://h/p')
    expect(js).toContain('startsWith("/__media://")')
  })

  it('maps the media path back onto the lazy-media reader', async () => {
    const roots = fixtureRoot()
    const res = await get(roots, '/__media://abc/media/image1.png')
    expect(res.status).toBe(200)
    expect([...res.body]).toEqual([1, 2, 3])
    expect(res.headers['content-type']).toBe('image/png')
  })

  it('maps markdown assets and the html preview through their handlers', async () => {
    const roots = fixtureRoot()
    const md = await get(roots, '/__mdasset:///notes/pic/ok.png')
    expect(String(md.body)).toContain('md-asset:///notes/pic/ok.png')
    const html = await get(roots, '/__htmlasset://local/notes/a/')
    expect(String(html.body)).toContain('html-asset://local/notes/a/')
    const preview = await get(roots, '/__preview://view-42/index.html')
    expect(String(preview.body)).toContain('<p>preview</p>')
    const gone = await get(roots, '/__preview://view-7/index.html')
    expect(String(gone.body)).toContain('preview gone')
  })

  it('rejects a request without the token and accepts the cookie', async () => {
    const roots = fixtureRoot()
    const denied = await get(roots, '/docs/', 'sekret')
    expect(denied.status).toBe(401)
    const viaQuery = await get(roots, '/docs/?token=sekret', 'sekret')
    expect(viaQuery.status).toBe(200)
    expect(viaQuery.headers['set-cookie']).toContain('genoffice_web_token=sekret')
    const viaCookie = await get(roots, '/docs/', 'sekret', 'genoffice_web_token=sekret')
    expect(viaCookie.status).toBe(200)
  })

  it('does not serve files outside the renderer root', async () => {
    const roots = fixtureRoot()
    const res = await get(roots, '/docs/../../etc/passwd')
    expect(res.status).toBe(404)
  })

  it('reports unbuilt renderers with a hint instead of crashing', async () => {
    const res = await get({}, '/slides/')
    expect(res.status).toBe(404)
    expect(String(res.body)).toContain('not built')
  })

  it('computes renderer roots from the monorepo layout in dev', () => {
    const roots = webRendererRoots()
    expect(roots.docs).toBe('/tmp/nowhere/apps/docs/out/renderer')
    expect(roots.html).toBe('/tmp/nowhere/apps/html/out/renderer')
  })
})

import { describe, expect, it } from 'vitest'
import {
  imageObjectKey,
  isImageHostUsable,
  normalizeImageHostConfig,
  sigV4,
  uploadImageToHost,
} from '../src/main/image-host'
import type { ImageHostConfig } from '../src/shared/ipc'

const fetchJson = (status: number, body: unknown) => async () => ({
  ok: status >= 200 && status < 300,
  status,
  headers: { get: () => null },
  text: async () => JSON.stringify(body),
})

describe('normalizeImageHostConfig', () => {
  it('accepts a complete S3 configuration and rejects a half-filled one', () => {
    const config = normalizeImageHostConfig({
      kind: 's3',
      endpoint: 'https://acc.r2.cloudflarestorage.com',
      bucket: 'notes',
      accessKeyId: 'AKID',
      secretAccessKey: 'shhh',
    })
    expect(config?.kind).toBe('s3')
    expect(config?.region).toBe('auto')
    expect(normalizeImageHostConfig({ kind: 's3', endpoint: 'https://x', bucket: 'b' })).toBeNull()
  })

  it('rejects non-https S3 endpoints, unknown kinds and missing tokens', () => {
    expect(
      normalizeImageHostConfig({
        kind: 's3',
        endpoint: 'http://acc.r2',
        bucket: 'b',
        accessKeyId: 'a',
        secretAccessKey: 's',
      }),
    ).toBeNull()
    expect(normalizeImageHostConfig({ kind: 'ftp' })).toBeNull()
    expect(normalizeImageHostConfig({ kind: 'smms' })).toBeNull()
    expect(normalizeImageHostConfig({ kind: 'github', token: 't' })).toBeNull()
    expect(normalizeImageHostConfig('nonsense')).toBeNull()
    expect(normalizeImageHostConfig(null)).toBeNull()
  })

  it('isImageHostUsable mirrors the normalization', () => {
    expect(isImageHostUsable({ kind: 'smms', token: 't' })).toBe(true)
    expect(isImageHostUsable({ kind: 'smms' })).toBe(false)
    expect(isImageHostUsable(null)).toBe(false)
  })
})

describe('sigV4', () => {
  // AWS documentation "Example: GET Object" — the SigV4 test vector for S3
  // itself (empty query, host/range/x-amz-content-sha256/x-amz-date headers).
  const url = new URL('https://examplebucket.s3.amazonaws.com/test.txt')
  const headers = {
    Host: 'examplebucket.s3.amazonaws.com',
    Range: 'bytes=0-9',
    'x-amz-content-sha256': 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855',
    'x-amz-date': '20130524T000000Z',
  }
  it('reproduces the documented S3 signature end to end', () => {
    const authorization = sigV4(
      'GET',
      url,
      headers,
      Buffer.from(''),
      {
        accessKeyId: 'AKIAIOSFODNN7EXAMPLE', // public-hygiene: fixture
        secretAccessKey: 'wJalrXUtnFEMI/K7MDENG/bPxRfiCYEXAMPLEKEY',
        region: 'us-east-1',
        service: 's3',
      },
      '20130524T000000Z',
    )
    expect(authorization).toContain(
      'Credential=AKIAIOSFODNN7EXAMPLE/20130524/us-east-1/s3/aws4_request', // public-hygiene: fixture
    )
    expect(authorization).toContain('SignedHeaders=host;range;x-amz-content-sha256;x-amz-date')
    expect(authorization).toContain(
      'Signature=f0e8bdb87c964420e857bd35b5d6ed310bd44f0170aba48dd91039c6036bdb41',
    )
  })
})

describe('uploadImageToHost — S3', () => {
  const config: ImageHostConfig = {
    kind: 's3',
    endpoint: 'https://acc.r2.cloudflarestorage.com',
    region: 'auto',
    bucket: 'notes',
    accessKeyId: 'AKID',
    secretAccessKey: 'shhh',
    prefix: 'md/',
    publicBase: 'https://cdn.example.com',
  }
  const now = new Date('2026-09-30T08:09:00Z')

  it('PUTs to endpoint/bucket/key and returns publicBase/key', async () => {
    const calls: Array<{
      url: string
      init?: { method?: string; headers?: Record<string, string> }
    }> = []
    const result = await uploadImageToHost(config, {
      bytes: Buffer.from('pngdata'),
      ext: 'png',
      now,
      fetchImpl: async (url, init) => {
        calls.push({ url, init })
        return { ok: true, status: 200, headers: { get: () => null }, text: async () => '' }
      },
    })
    expect(result.ok).toBe(true)
    expect(result.url).toMatch(/^https:\/\/cdn\.example\.com\/md\/20260930\/\w+\.png$/)
    expect(calls[0]!.init?.method).toBe('PUT')
    expect(calls[0]!.url).toContain('https://acc.r2.cloudflarestorage.com/notes/md/')
    // the signed headers chain covers host + the two x-amz headers
    expect(calls[0]!.init?.headers?.Authorization).toContain(
      'SignedHeaders=host;x-amz-content-sha256;x-amz-date',
    )
  })

  it('without publicBase the API endpoint is the URL base; a 403 surfaces as a failure', async () => {
    const noBase = { ...config, publicBase: undefined }
    const ok = await uploadImageToHost(noBase, {
      bytes: Buffer.from('x'),
      ext: 'png',
      now,
      fetchImpl: async () => ({
        ok: true,
        status: 200,
        headers: { get: () => null },
        text: async () => '',
      }),
    })
    expect(ok.url).toMatch(/^https:\/\/acc\.r2\.cloudflarestorage\.com\/notes\/md\//)
    const fail = await uploadImageToHost(noBase, {
      bytes: Buffer.from('x'),
      ext: 'png',
      now,
      fetchImpl: async () => ({
        ok: false,
        status: 403,
        headers: { get: () => null },
        text: async () => '<xml>denied</xml>',
      }),
    })
    expect(fail.ok).toBe(false)
    expect(fail.error).toContain('403')
  })
})

describe('uploadImageToHost — SM.MS and GitHub', () => {
  const now = new Date('2026-09-30T08:09:00Z')

  it('posts the file with the raw API token and reads data.url', async () => {
    let seen = ''
    const result = await uploadImageToHost(
      { kind: 'smms', token: 'tok' },
      {
        bytes: Buffer.from('gifdata'),
        ext: 'gif',
        now,
        fetchImpl: async (url, init) => {
          seen = `${url} ${String(init?.headers?.Authorization)}`
          return fetchJson(200, { success: true, data: { url: 'https://i.smms.app/x.gif' } })()
        },
      },
    )
    expect(result).toEqual({ ok: true, url: 'https://i.smms.app/x.gif' })
    expect(seen).toContain('https://smms.app/api/v2/upload')
    expect(seen).toContain('tok')
    expect(seen).not.toContain('Bearer')
  })

  it('an SM.MS rejection carries the host message', async () => {
    const result = await uploadImageToHost(
      { kind: 'smms', token: 'tok' },
      {
        bytes: Buffer.from('x'),
        ext: 'png',
        now,
        fetchImpl: fetchJson(200, {
          success: false,
          message: 'image upload repeated limit exceeded',
        }),
      },
    )
    expect(result.ok).toBe(false)
    expect(result.error).toContain('image upload repeated')
  })

  it('commits to the contents API and prefers the configured CDN base', async () => {
    const bodies: string[] = []
    const config: ImageHostConfig = {
      kind: 'github',
      token: 'ghp',
      owner: 'user',
      repo: 'notes',
      branch: 'media',
      dir: 'img/',
      urlPrefix: 'https://cdn.jsdelivr.net/gh/user/notes@media',
    }
    const result = await uploadImageToHost(config, {
      bytes: Buffer.from('jpgdata'),
      ext: 'jpg',
      now,
      fetchImpl: async (url, init) => {
        bodies.push(String(init?.body))
        return fetchJson(201, {
          content: { download_url: 'https://raw.githubusercontent.com/user/notes/media/img/x.jpg' },
        })()
      },
    })
    expect(result.ok).toBe(true)
    expect(result.url).toMatch(
      /^https:\/\/cdn\.jsdelivr\.net\/gh\/user\/notes@media\/img\/\d{8}\/\w+\.jpg$/,
    )
    const sent = JSON.parse(bodies[0]!) as { branch: string; content: string }
    expect(sent.branch).toBe('media')
    expect(Buffer.from(sent.content, 'base64').toString()).toBe('jpgdata')
  })

  it('an unusable configuration never issues a request', async () => {
    let called = false
    const result = await uploadImageToHost(null, {
      bytes: Buffer.from('x'),
      ext: 'png',
      now,
      fetchImpl: async () => {
        called = true
        return fetchJson(200, {})()
      },
    })
    expect(result.ok).toBe(false)
    expect(called).toBe(false)
  })
})

describe('imageObjectKey', () => {
  it('date-shards under the prefix with the requested extension', () => {
    const key = imageObjectKey('notes/', 'png', new Date('2026-09-30T08:09:00Z'))
    expect(key).toMatch(/^notes\/20260930\/[a-z0-9]+\.png$/)
    expect(imageObjectKey(undefined, 'jpg', new Date('2026-09-30T08:09:00Z'))).toMatch(
      /^20260930\//,
    )
  })
})

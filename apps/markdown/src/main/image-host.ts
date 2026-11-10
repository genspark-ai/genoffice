// Bring-your-own image host for Markdown pastes/drops (genoffice#388): the
// maintainer-accepted shape is "configure your own storage, keep the local
// assets/ copy as the fallback". Three adapters cover the hosts the issue
// names — S3-compatible (Cloudflare R2 / Aliyun OSS / Tencent COS / MinIO),
// SM.MS, GitHub — and all run in the main process, so no renderer CORS
// applies and the credentials never enter a web context.
import { createHash, createHmac } from 'node:crypto'
import type { ImageHostConfig, ImageUploadResult, ImageHostKind } from '../shared/ipc'

const UPLOAD_TIMEOUT_MS = 20_000
/** one image upload, not a transfer service: a hard ceiling, well above the paste cap */
const MAX_UPLOAD_BYTES = 25 * 1024 * 1024
/** responses beyond this are not JSON these hosts produce */
const MAX_RESPONSE_BYTES = 1024 * 1024

const sha256Hex = (data: Buffer | string): string => createHash('sha256').update(data).digest('hex')
const hmac = (key: Buffer | string, data: string): Buffer =>
  createHmac('sha256', key).update(data).digest()

type FetchLike = (
  url: string,
  init?: {
    method?: string
    headers?: Record<string, string>
    body?: unknown
    signal?: AbortSignal
  },
) => Promise<{
  ok: boolean
  status: number
  headers?: { get?: (name: string) => string | null }
  text?: () => Promise<string>
}>

/** the config file shape is what the dialog hands over; anything malformed disables the host */
export function normalizeImageHostConfig(raw: unknown): ImageHostConfig | null {
  if (raw === null || raw === undefined) return null
  if (typeof raw !== 'object') return null
  const c = raw as Record<string, unknown>
  const kind = c.kind === 's3' || c.kind === 'smms' || c.kind === 'github' ? c.kind : null
  if (!kind) return null
  const str = (v: unknown, max = 512): string =>
    typeof v === 'string' ? v.trim().slice(0, max) : ''
  const config: ImageHostConfig = { kind }
  if (kind === 's3') {
    config.endpoint = str(c.endpoint)
    config.region = str(c.region, 64) || 'auto'
    config.bucket = str(c.bucket, 128)
    config.accessKeyId = str(c.accessKeyId, 256)
    config.secretAccessKey = str(c.secretAccessKey, 256)
    config.prefix = str(c.prefix, 256)
    config.publicBase = str(c.publicBase)
    if (!config.endpoint || !config.bucket || !config.accessKeyId || !config.secretAccessKey)
      return null
    if (!/^https:\/\//i.test(config.endpoint)) return null
  } else if (kind === 'smms') {
    config.token = str(c.token, 256)
    if (!config.token) return null
  } else {
    config.token = str(c.token, 256)
    config.owner = str(c.owner, 128)
    config.repo = str(c.repo, 128)
    config.branch = str(c.branch, 128) || 'main'
    config.dir = str(c.dir, 256) || 'images/'
    config.urlPrefix = str(c.urlPrefix)
    if (!config.token || !config.owner || !config.repo) return null
  }
  return config
}

export function isImageHostUsable(config: ImageHostConfig | null | undefined): boolean {
  return normalizeImageHostConfig(config) !== null
}

/** object key inside the host, date-sharded so a directory listing stays navigable */
export function imageObjectKey(prefix: string | undefined, ext: string, now: Date): string {
  const date = now.toISOString().slice(0, 10).replace(/-/g, '')
  const id = `${now.getTime().toString(36)}${Math.floor(Math.random() * 0xffff).toString(36)}`
  return `${(prefix ?? '').replace(/^\/+|\/+$/g, '')}${prefix ? '/' : ''}${date}/${id}.${ext}`
}

const trimSlash = (s: string): string => s.replace(/\/+$/, '')

/**
 * AWS Signature Version 4, path-style PUT. Written against the documented
 * algorithm (the iam.amazonaws.com ListUsers test vector pins the whole
 * chain); no payload profile beyond the signed SHA-256, which every
 * S3-compatible store accepts.
 */
export function sigV4(
  method: string,
  url: URL,
  headers: Record<string, string>,
  payload: Buffer,
  creds: { accessKeyId: string; secretAccessKey: string; region: string; service: string },
  amzDate: string,
): string {
  const dateStamp = amzDate.slice(0, 8)
  const payloadHash = sha256Hex(payload)
  const canonicalHeaders = Object.keys(headers)
    .map((h) => h.toLowerCase())
    .sort()
    .map(
      (h) => `${h}:${headers[Object.keys(headers).find((k) => k.toLowerCase() === h)!]!.trim()}\n`,
    )
    .join('')
  const signedHeaders = Object.keys(headers)
    .map((h) => h.toLowerCase())
    .sort()
    .join(';')
  const canonicalUri = url.pathname.split('/').map(encodeURIComponent).join('/') || '/'
  const canonicalQuery = [...url.searchParams.entries()]
    .map(([k, v]) => [
      encodeURIComponent(k).replace(/%20/g, '+'),
      encodeURIComponent(v).replace(/%20/g, '+'),
    ])
    .sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] < b[1] ? -1 : 1))
    .map(([k, v]) => `${k}=${v}`)
    .join('&')
  const canonicalRequest = [
    method.toUpperCase(),
    canonicalUri,
    canonicalQuery,
    canonicalHeaders,
    signedHeaders,
    payloadHash,
  ].join('\n')
  const stringToSign = [
    'AWS4-HMAC-SHA256',
    amzDate,
    `${dateStamp}/${creds.region}/${creds.service}/aws4_request`,
    sha256Hex(canonicalRequest),
  ].join('\n')
  const kDate = hmac(`AWS4${creds.secretAccessKey}`, dateStamp)
  const kRegion = hmac(kDate, creds.region)
  const kService = hmac(kRegion, creds.service)
  const kSigning = hmac(kService, 'aws4_request')
  const signature = createHmac('sha256', kSigning).update(stringToSign).digest('hex')
  return `AWS4-HMAC-SHA256 Credential=${creds.accessKeyId}/${dateStamp}/${creds.region}/${creds.service}/aws4_request, SignedHeaders=${signedHeaders}, Signature=${signature}`
}

async function readCapped(response: {
  headers?: { get?: (n: string) => string | null }
  text?: () => Promise<string>
}): Promise<string> {
  const declared = Number(response.headers?.get?.('content-length') ?? '0')
  if (Number.isFinite(declared) && declared > MAX_RESPONSE_BYTES)
    throw new Error('response too large for an upload result')
  const text = (await response.text?.()) ?? ''
  if (text.length > MAX_RESPONSE_BYTES) throw new Error('response too large for an upload result')
  return text
}

/** extract a JSON field by dotted path without pulling in a runtime dependency */
function jsonPath(json: unknown, path: string): string | undefined {
  let cur: unknown = json
  for (const seg of path.split('.')) {
    if (cur === null || typeof cur !== 'object') return undefined
    cur = (cur as Record<string, unknown>)[seg]
  }
  return typeof cur === 'string' ? cur : undefined
}

async function uploadS3(
  config: ImageHostConfig,
  bytes: Buffer,
  ext: string,
  now: Date,
  fetchImpl: FetchLike,
): Promise<ImageUploadResult> {
  const key = imageObjectKey(config.prefix, ext, now)
  const endpoint = trimSlash(config.endpoint!)
  const url = new URL(`${endpoint}/${config.bucket}/${key}`)
  const amzDate = now
    .toISOString()
    .replace(/[-:]/g, '')
    .replace(/\.\d{3}/, '')
  const headers: Record<string, string> = {
    Host: url.host,
    'x-amz-date': amzDate,
    'x-amz-content-sha256': sha256Hex(bytes),
  }
  headers.Authorization = sigV4(
    'PUT',
    url,
    headers,
    bytes,
    {
      accessKeyId: config.accessKeyId!,
      secretAccessKey: config.secretAccessKey!,
      region: config.region || 'auto',
      service: 's3',
    },
    amzDate,
  )
  const response = await fetchImpl(url.href, { method: 'PUT', headers, body: bytes })
  if (!response.ok) {
    const body = await readCapped(response).catch(() => '')
    return { ok: false, error: `S3 ${response.status}: ${body.slice(0, 200) || 'upload rejected'}` }
  }
  const base = config.publicBase ? trimSlash(config.publicBase) : `${endpoint}/${config.bucket}`
  return { ok: true, url: `${base}/${key}` }
}

async function uploadSmms(
  config: ImageHostConfig,
  bytes: Buffer,
  ext: string,
  _now: Date,
  fetchImpl: FetchLike,
): Promise<ImageUploadResult> {
  const form = new FormData()
  form.append(
    'smfile',
    new Blob([new Uint8Array(bytes)], { type: `image/${ext === 'jpg' ? 'jpeg' : ext}` }),
    `image.${ext}`,
  )
  const response = await fetchImpl('https://smms.app/api/v2/upload', {
    method: 'POST',
    headers: { Authorization: config.token ?? '' },
    body: form,
  })
  const text = await readCapped(response).catch(() => '')
  let json: unknown
  try {
    json = JSON.parse(text)
  } catch {
    return { ok: false, error: `SM.MS ${response.status}: unparseable response` }
  }
  if (jsonPath(json, 'success') !== 'true' && (json as { success?: unknown })?.success !== true)
    return {
      ok: false,
      error: `SM.MS: ${String(jsonPath(json, 'message') ?? 'upload rejected').slice(0, 200)}`,
    }
  const url = jsonPath(json, 'data.url') ?? jsonPath(json, 'data.quilt_url')
  if (!url) return { ok: false, error: 'SM.MS: no URL in response' }
  return { ok: true, url }
}

async function uploadGithub(
  config: ImageHostConfig,
  bytes: Buffer,
  ext: string,
  now: Date,
  fetchImpl: FetchLike,
): Promise<ImageUploadResult> {
  const path = imageObjectKey(config.dir, ext, now)
  const branch = config.branch || 'main'
  const url = new URL(
    `https://api.github.com/repos/${config.owner}/${config.repo}/contents/${path}`,
  )
  const response = await fetchImpl(url.href, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${config.token}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      message: `image ${now.toISOString()}`,
      content: bytes.toString('base64'),
      branch,
    }),
  })
  const text = await readCapped(response).catch(() => '')
  let json: unknown
  try {
    json = JSON.parse(text)
  } catch {
    return { ok: false, error: `GitHub ${response.status}: unparseable response` }
  }
  if (!response.ok)
    return {
      ok: false,
      error: `GitHub ${response.status}: ${String(jsonPath(json, 'message') ?? 'upload rejected').slice(0, 200)}`,
    }
  const download = jsonPath(json, 'content.download_url')
  if (!download) return { ok: false, error: 'GitHub: no download_url in response' }
  // a configured CDN base (jsDelivr etc.) replaces the API's raw URL; the
  // repo-relative path is identical
  const url2 = config.urlPrefix ? `${trimSlash(config.urlPrefix)}/${path}` : download
  return { ok: true, url: url2 }
}

export async function uploadImageToHost(
  config: ImageHostConfig | null,
  input: { bytes: Buffer; ext: string; fetchImpl?: FetchLike; now?: Date },
): Promise<ImageUploadResult> {
  const normalized = normalizeImageHostConfig(config)
  if (!normalized) return { ok: false, error: 'no usable image host configured' }
  if (input.bytes.length === 0) return { ok: false, error: 'empty image' }
  if (input.bytes.length > MAX_UPLOAD_BYTES)
    return { ok: false, error: `image exceeds ${Math.round(MAX_UPLOAD_BYTES / 1024 / 1024)} MB` }
  const baseFetch = input.fetchImpl ?? (fetch as unknown as FetchLike)
  // a wedged host must not hang the paste path: every adapter call carries the
  // same abort deadline (covers connect, body streaming and response read)
  const fetchImpl: FetchLike = (url, init) =>
    baseFetch(url, { ...init, signal: AbortSignal.timeout(UPLOAD_TIMEOUT_MS) })
  const now = input.now ?? new Date()
  const run =
    normalized.kind === 's3' ? uploadS3 : normalized.kind === 'smms' ? uploadSmms : uploadGithub
  try {
    return await run(normalized, input.bytes, input.ext, now, fetchImpl)
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : 'upload failed',
    }
  }
}

export type { ImageHostKind }

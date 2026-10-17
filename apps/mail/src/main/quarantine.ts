import {
  closeSync,
  copyFileSync,
  existsSync,
  mkdirSync,
  openSync,
  readSync,
  rmSync,
  statSync,
} from 'node:fs'
import { join, resolve, sep } from 'node:path'
import {
  GMAIL_MAX_MESSAGE_BYTES,
  GMAIL_MAX_RECIPIENTS_PER_MESSAGE,
  GMAIL_WARN_BYTES,
  GRAPH_MAX_ATTACHMENT_BYTES,
  GRAPH_MAX_RECIPIENTS,
  GRAPH_MAX_TOTAL_BYTES,
  GRAPH_WARN_TOTAL_BYTES,
  INLINE_PREVIEW_CAP_BYTES,
  MAX_ATTACHMENT_BYTES,
  MAX_RECIPIENTS,
  MAX_TOTAL_BYTES,
  OUTGOING_WARN_BYTES,
  contentAddress,
  safeFilename,
} from '@genoffice/mail-protocols'
import type { StageInput, StagedAttachment } from '@genoffice/mail-protocols'
import type { ParsedFile } from '@genoffice/file-parse'

// Attachment quarantine pipeline (H5 owns this file). Normative order per
// the H1 spec section 3: staged temp --> quarantine scan --> store. The
// scan runs in main only; hostile-file parsing never runs in the renderer.
// Preview text extraction delegates to the battle-tested file-parse path
// (same extractor the docs pipeline uses); quarantine never invents a
// second parser. Zero new external prod deps: file-parse is already a
// monorepo workspace package (docs/sheets/html precedent).

/** Scan outcome: clean previews, quarantined stores restricted, blocked drops. */
export type QuarantineVerdict = 'clean' | 'quarantined' | 'blocked'

export type QuarantineReason =
  | 'clean'
  | 'traversal'
  | 'hostile-name'
  | 'exec-type'
  | 'script-type'
  | 'macro-type'
  | 'mime-mismatch'
  | 'over-cap'
  | 'suspicious-content'

export type MailDriverKind = 'imap' | 'gmail' | 'graph'

export interface SendCaps {
  warnTotalBytes: number
  maxTotalBytes: number
  maxAttachmentBytes: number
  maxRecipients: number
}

/** Content sniffing reads at most this many leading bytes per file. */
export const QUARANTINE_SCAN_BYTES = 1024 * 1024

/**
 * EICAR-style tripwire WITHOUT the live signature: the standard test
 * vector starts with this prefix and is exactly 68 bytes. Matching only
 * the prefix keeps the real signature out of the tree (and out of the
 * reach of desktop AV scanners) while still pinning the detection path.
 */
const EICAR_PREFIX = 'X5O!P%@AP[4'
const EICAR_LENGTH = 68

/** Zip-bomb tripwires for the bounded local-header scan below. */
const ZIP_MAX_ENTRIES = 10_000
const ZIP_MAX_UNCOMPRESSED_BYTES = 1024 * 1024 * 1024

/** Executables are NEVER inline-previewed: save + OS-open only. */
const EXECUTABLE_EXTS = new Set([
  'exe',
  'msi',
  'com',
  'scr',
  'pif',
  'cpl',
  'dll',
  'sys',
  'ocx',
  'jar',
  'apk',
  'app',
])

/** Scripts are NEVER inline-previewed: save + OS-open only. */
const SCRIPT_EXTS = new Set([
  'bat',
  'cmd',
  'ps1',
  'psm1',
  'vbs',
  'vbe',
  'wsf',
  'wsh',
  'sh',
  'bash',
  'py',
  'rb',
  'pl',
  'php',
  'lua',
  'url',
  'lnk',
  'reg',
])

/** Macro-enabled Office formats are NEVER inline-previewed. */
const MACRO_EXTS = new Set(['docm', 'dotm', 'xlsm', 'xltm', 'pptm', 'potm', 'ppam', 'sldm'])

/** Common document extensions used for the double-extension tripwire. */
const DOCUMENT_EXTS = new Set([
  'pdf',
  'doc',
  'docx',
  'xls',
  'xlsx',
  'ppt',
  'pptx',
  'txt',
  'md',
  'csv',
  'png',
  'jpg',
  'jpeg',
  'gif',
  'webp',
  'mp3',
  'mp4',
])

/** Expected MIME per extension; octet-stream/empty never count as mismatch. */
const EXT_MIME: Record<string, string> = {
  txt: 'text/plain',
  log: 'text/plain',
  md: 'text/markdown',
  csv: 'text/csv',
  json: 'application/json',
  xml: 'text/xml',
  html: 'text/html',
  pdf: 'application/pdf',
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
  doc: 'application/msword',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xls: 'application/vnd.ms-excel',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ppt: 'application/vnd.ms-powerpoint',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  zip: 'application/zip',
  mp3: 'audio/mpeg',
  mp4: 'video/mp4',
}

/** Control characters (including NUL) are rejected from attachment names. */
function hasControlChars(value: string): boolean {
  for (let i = 0; i < value.length; i += 1) {
    const code = value.charCodeAt(i)
    if (code < 32 || code === 127) return true
  }
  return false
}

export interface QuarantineScanResult {
  verdict: QuarantineVerdict
  reason: QuarantineReason
  /** Sanitized bare filename (safeFilename form), safe to join under a root. */
  filename: string
  /** True when the bytes exceed the warn threshold but stay under the cap. */
  sizeWarn: boolean
  detail?: string
}

/**
 * Per-driver send caps resolved where the driver is known (Q-5). Unknown
 * drivers fall back to the IMAP/SMTP mirror (the compose-send default).
 * Gmail numbers are raw-byte pre-checks; base64 inflation is enforced by
 * the driver's checkSendLimits at send time.
 */
export function resolveSendCaps(driver?: MailDriverKind): SendCaps {
  if (driver === 'gmail') {
    return {
      warnTotalBytes: GMAIL_WARN_BYTES,
      maxTotalBytes: GMAIL_MAX_MESSAGE_BYTES,
      maxAttachmentBytes: GMAIL_MAX_MESSAGE_BYTES,
      maxRecipients: GMAIL_MAX_RECIPIENTS_PER_MESSAGE,
    }
  }
  if (driver === 'graph') {
    return {
      warnTotalBytes: GRAPH_WARN_TOTAL_BYTES,
      maxTotalBytes: GRAPH_MAX_TOTAL_BYTES,
      maxAttachmentBytes: GRAPH_MAX_ATTACHMENT_BYTES,
      maxRecipients: GRAPH_MAX_RECIPIENTS,
    }
  }
  return {
    warnTotalBytes: OUTGOING_WARN_BYTES,
    maxTotalBytes: MAX_TOTAL_BYTES,
    maxAttachmentBytes: MAX_ATTACHMENT_BYTES,
    maxRecipients: MAX_RECIPIENTS,
  }
}

/**
 * Staged-path confinement (M1 precedent, same rule): the candidate must
 * resolve to the root itself or inside it. Throws a mail-prefixed error
 * otherwise. Applies to stageDir, quarantineDir, and storeDir alike (Q-1).
 */
export function confinePath(candidate: string, root: string): string {
  const base = resolve(root)
  const confined = resolve(candidate)
  if (confined !== base && !confined.startsWith(base + sep)) {
    throw new Error('mail: path outside quarantine root')
  }
  return confined
}

function extOf(filename: string): string {
  const base = filename.split('.').pop() ?? ''
  return base.toLowerCase()
}

/**
 * Filename scan on the ORIGINAL renderer/supplier-supplied name (before
 * safeFilename stripping, so traversal and ADS tricks stay visible).
 * Renderer-supplied names must be bare: any separator is traversal.
 */
export function scanFileName(raw: string): {
  ok: boolean
  reason?: QuarantineReason
  detail?: string
} {
  if (raw.length === 0 || raw.length > 256) {
    return { ok: false, reason: 'hostile-name', detail: 'empty or overlong name' }
  }
  if (hasControlChars(raw)) {
    return { ok: false, reason: 'hostile-name', detail: 'control characters' }
  }
  if (
    raw.includes('/') ||
    raw.includes('\\') ||
    raw === '..' ||
    /^[a-zA-Z]:/.test(raw) ||
    raw.startsWith('\\\\')
  ) {
    return { ok: false, reason: 'traversal', detail: 'path separators or drive/UNC form' }
  }
  const segments = raw.split(/[\\/]/)
  if (segments.some((part) => part === '..')) {
    return { ok: false, reason: 'traversal', detail: 'parent segment' }
  }
  if (raw.includes(':')) {
    return { ok: false, reason: 'hostile-name', detail: 'ADS/colon trick' }
  }
  if (/[. ]$/.test(raw)) {
    return { ok: false, reason: 'hostile-name', detail: 'trailing dot/space (Windows strip trick)' }
  }
  const safe = safeFilename(raw)
  if (safe === 'unnamed' && raw.trim().replace(/^\.+/, '') === '') {
    return { ok: false, reason: 'hostile-name', detail: 'nothing left after sanitize' }
  }
  return { ok: true }
}

function neverInlineReason(safeName: string): QuarantineReason | null {
  const ext = extOf(safeName)
  if (EXECUTABLE_EXTS.has(ext)) return 'exec-type'
  if (SCRIPT_EXTS.has(ext)) return 'script-type'
  if (MACRO_EXTS.has(ext)) return 'macro-type'
  const parts = safeName.toLowerCase().split('.')
  if (
    parts.length >= 3 &&
    DOCUMENT_EXTS.has(parts[parts.length - 2] ?? '') &&
    (EXECUTABLE_EXTS.has(ext) || SCRIPT_EXTS.has(ext))
  ) {
    return 'exec-type'
  }
  return null
}

function mimeMismatch(safeName: string, mime: string): boolean {
  const normalized = mime.trim().toLowerCase().split(';')[0]?.trim() ?? ''
  if (normalized === '' || normalized === 'application/octet-stream') return false
  const expected = EXT_MIME[extOf(safeName)]
  if (!expected) return false
  if (normalized === expected) return false
  if (expected === 'text/xml' && normalized === 'application/xml') return false
  return true
}

/**
 * Bounded content heuristics over the leading prefix (at most
 * QUARANTINE_SCAN_BYTES). No inflation, no full parse: the zip scan walks
 * local headers only and stops at the entry/byte tripwires.
 */
export function scanContentPrefix(
  safeName: string,
  prefix: Uint8Array,
  fullSize: number,
): { verdict: QuarantineVerdict; reason: QuarantineReason; detail?: string } {
  const text = String.fromCharCode(...prefix.subarray(0, Math.min(prefix.length, 64)))
  if (text.startsWith(EICAR_PREFIX) && fullSize === EICAR_LENGTH) {
    return { verdict: 'blocked', reason: 'suspicious-content', detail: 'EICAR-style test vector' }
  }
  const ext = extOf(safeName)
  const isMz = prefix.length >= 2 && prefix[0] === 0x4d && prefix[1] === 0x5a
  const isElf =
    prefix.length >= 4 &&
    prefix[0] === 0x7f &&
    prefix[1] === 0x45 &&
    prefix[2] === 0x4c &&
    prefix[3] === 0x46
  if ((isMz || isElf) && !EXECUTABLE_EXTS.has(ext)) {
    return {
      verdict: 'blocked',
      reason: 'suspicious-content',
      detail: 'executable magic under document name',
    }
  }
  const looksZip =
    prefix.length >= 4 &&
    prefix[0] === 0x50 &&
    prefix[1] === 0x4b &&
    prefix[2] === 0x03 &&
    prefix[3] === 0x04
  if (looksZip) {
    const bomb = scanZipHeaders(prefix)
    if (bomb) return { verdict: 'blocked', reason: 'suspicious-content', detail: bomb }
  }
  if (ext === 'pdf' && !(prefix.length >= 5 && text.startsWith('%PDF-'))) {
    return { verdict: 'quarantined', reason: 'mime-mismatch', detail: 'pdf name without pdf magic' }
  }
  return { verdict: 'clean', reason: 'clean' }
}

/** Walk zip local headers (no inflation); returns a detail string when bomb-like. */
function scanZipHeaders(prefix: Uint8Array): string | null {
  let offset = 0
  let entries = 0
  let uncompressed = 0
  while (offset + 30 <= prefix.length) {
    if (
      prefix[offset] !== 0x50 ||
      prefix[offset + 1] !== 0x4b ||
      prefix[offset + 2] !== 0x03 ||
      prefix[offset + 3] !== 0x04
    ) {
      break
    }
    const view = new DataView(prefix.buffer, prefix.byteOffset + offset, 30)
    const compressed = view.getUint32(18, true)
    const uncompressedSize = view.getUint32(22, true)
    const nameLen = view.getUint16(26, true)
    const extraLen = view.getUint16(28, true)
    entries += 1
    uncompressed += uncompressedSize
    if (entries > ZIP_MAX_ENTRIES) return 'zip entry count over tripwire'
    if (uncompressed > ZIP_MAX_UNCOMPRESSED_BYTES) return 'zip uncompressed size over tripwire'
    if (compressed > prefix.length) break
    offset += 30 + nameLen + extraLen + compressed
    if (offset <= 30) break
  }
  return null
}

/** Read at most QUARANTINE_SCAN_BYTES from the head of a file. */
export function readPrefix(path: string, maxBytes: number = QUARANTINE_SCAN_BYTES): Uint8Array {
  const fd = openSync(path, 'r')
  try {
    const size = Math.min(statSync(path).size, maxBytes)
    const buf = Buffer.alloc(size)
    readSync(fd, buf, 0, size, 0)
    return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength)
  } finally {
    closeSync(fd)
  }
}

export interface ScanBytesInput {
  filename: string
  mime: string
  bytes: Uint8Array
}

/**
 * Full scan over in-memory bytes (compose pre-send surfacing): name, size
 * cap, never-inline families, MIME/extension mismatch, content heuristics.
 */
export function scanAttachmentBytes(input: ScanBytesInput, caps: SendCaps): QuarantineScanResult {
  const nameScan = scanFileName(input.filename)
  const safe = safeFilename(input.filename)
  if (!nameScan.ok) {
    return {
      verdict: 'blocked',
      reason: nameScan.reason ?? 'hostile-name',
      filename: safe,
      sizeWarn: false,
      detail: nameScan.detail,
    }
  }
  if (
    input.bytes.byteLength > caps.maxAttachmentBytes ||
    input.bytes.byteLength > INLINE_PREVIEW_CAP_BYTES
  ) {
    return { verdict: 'blocked', reason: 'over-cap', filename: safe, sizeWarn: false }
  }
  const inline = neverInlineReason(safe)
  if (inline) {
    return {
      verdict: 'quarantined',
      reason: inline,
      filename: safe,
      sizeWarn: input.bytes.byteLength > caps.warnTotalBytes,
    }
  }
  if (mimeMismatch(safe, input.mime)) {
    return {
      verdict: 'quarantined',
      reason: 'mime-mismatch',
      filename: safe,
      sizeWarn: input.bytes.byteLength > caps.warnTotalBytes,
    }
  }
  const prefix = input.bytes.subarray(0, Math.min(input.bytes.byteLength, QUARANTINE_SCAN_BYTES))
  const content = scanContentPrefix(safe, prefix, input.bytes.byteLength)
  return {
    verdict: content.verdict,
    reason: content.reason,
    filename: safe,
    sizeWarn: input.bytes.byteLength > caps.warnTotalBytes,
    detail: content.detail,
  }
}

/**
 * Scan for a file already on disk (compose-send enforcement, receive
 * ingest): same verdicts, prefix read bounded, never full-file parse.
 */
export function scanAttachmentPath(
  storedPath: string,
  filename: string,
  mime: string,
  size: number,
  caps: SendCaps,
): QuarantineScanResult {
  const nameScan = scanFileName(filename)
  const safe = safeFilename(filename)
  if (!nameScan.ok) {
    return {
      verdict: 'blocked',
      reason: nameScan.reason ?? 'hostile-name',
      filename: safe,
      sizeWarn: false,
      detail: nameScan.detail,
    }
  }
  if (size > caps.maxAttachmentBytes || size > INLINE_PREVIEW_CAP_BYTES) {
    return { verdict: 'blocked', reason: 'over-cap', filename: safe, sizeWarn: false }
  }
  const inline = neverInlineReason(safe)
  if (inline) {
    return {
      verdict: 'quarantined',
      reason: inline,
      filename: safe,
      sizeWarn: size > caps.warnTotalBytes,
    }
  }
  if (mimeMismatch(safe, mime)) {
    return {
      verdict: 'quarantined',
      reason: 'mime-mismatch',
      filename: safe,
      sizeWarn: size > caps.warnTotalBytes,
    }
  }
  const content = scanContentPrefix(safe, readPrefix(storedPath), size)
  return {
    verdict: content.verdict,
    reason: content.reason,
    filename: safe,
    sizeWarn: size > caps.warnTotalBytes,
    detail: content.detail,
  }
}

export interface PipelineDirs {
  stageDir: string
  quarantineDir: string
  storeDir: string
}

export type PipelineProgress = 'staged' | 'scanning' | 'stored' | 'blocked'

export interface PipelineInput {
  filename: string
  mime: string
  bytes: Uint8Array
}

export interface PipelineOutcome extends QuarantineScanResult {
  /** Disk home for clean (storeDir) or quarantined (quarantineDir); absent when blocked. */
  path?: string
}

export interface PipelineDeps {
  /** Canonical staging (stageAttachment from mail-protocols by default). */
  stage: (input: StageInput) => StagedAttachment
  onProgress?: (state: PipelineProgress) => void
}

/**
 * Staged-temp --> quarantine-scan --> store. Bytes land content-addressed
 * under stageDir first (normative order), then the scan promotes clean
 * files to storeDir, parks executables/mismatches in quarantineDir, and
 * deletes blocked files outright. Blocked outcomes carry no path.
 */
export function runAttachmentPipeline(
  input: PipelineInput,
  dirs: PipelineDirs,
  deps: PipelineDeps,
  caps: SendCaps = resolveSendCaps('imap'),
): PipelineOutcome {
  mkdirSync(dirs.stageDir, { recursive: true })
  mkdirSync(dirs.quarantineDir, { recursive: true })
  mkdirSync(dirs.storeDir, { recursive: true })
  const staged = deps.stage({
    filename: input.filename,
    mime: input.mime,
    bytes: input.bytes,
    stageDir: dirs.stageDir,
  })
  confinePath(staged.path, dirs.stageDir)
  deps.onProgress?.('staged')
  const scanned = scanAttachmentPath(
    staged.path,
    input.filename,
    input.mime,
    input.bytes.byteLength,
    caps,
  )
  deps.onProgress?.('scanning')
  if (scanned.verdict === 'blocked') {
    rmSync(staged.path, { force: true })
    deps.onProgress?.('blocked')
    return scanned
  }
  const digest = contentAddress(input.bytes)
  const home = scanned.verdict === 'clean' ? dirs.storeDir : dirs.quarantineDir
  const dest = confinePath(join(home, `${digest}-${scanned.filename}`), home)
  copyFileSync(staged.path, dest)
  rmSync(staged.path, { force: true })
  deps.onProgress?.('stored')
  return { ...scanned, path: dest }
}

export type PreviewExtractor = (filePath: string) => Promise<ParsedFile>

/** Lazy production extractor: the shared file-parse path, loaded on demand. */
async function defaultExtractor(filePath: string): Promise<ParsedFile> {
  const { parseFileToText } = await import('@genoffice/file-parse')
  return parseFileToText(filePath)
}

export interface PreviewOutcome {
  allowed: boolean
  reason: QuarantineReason
  text?: string
  kind?: ParsedFile['kind']
  detail?: string
}

/**
 * Executables-never-inline gate (Q-4): only clean-verdict files with
 * non-executable extensions reach the file-parse extractor. The extractor
 * is injected in tests; production lazily uses the shared file-parse path.
 */
export async function extractPreviewText(
  storedPath: string,
  filename: string,
  storeDir: string,
  verdict: QuarantineVerdict,
  extractor: PreviewExtractor = defaultExtractor,
): Promise<PreviewOutcome> {
  confinePath(storedPath, storeDir)
  if (!existsSync(storedPath)) {
    return { allowed: false, reason: 'hostile-name', detail: 'missing stored file' }
  }
  if (verdict !== 'clean') {
    return { allowed: false, reason: verdict === 'blocked' ? 'suspicious-content' : 'exec-type' }
  }
  const inline = neverInlineReason(safeFilename(filename))
  if (inline) {
    return { allowed: false, reason: inline }
  }
  const parsed = await extractor(storedPath)
  if (!parsed.ok) {
    return { allowed: false, reason: 'mime-mismatch', detail: parsed.error }
  }
  return { allowed: true, reason: 'clean', text: parsed.text, kind: parsed.kind }
}

/** Disk roots for one pipeline run (tests use temp dirs; prod wires app dirs). */
export function defaultPipelineDirs(root: string): PipelineDirs {
  return {
    stageDir: join(root, 'stage'),
    quarantineDir: join(root, 'quarantine'),
    storeDir: join(root, 'store'),
  }
}

import { createHash } from 'node:crypto'
import { constants } from 'node:fs'
import {
  copyFile,
  lstat,
  mkdir,
  readFile,
  readdir,
  realpath,
  rm,
  stat,
  writeFile,
} from 'node:fs/promises'
import { basename, dirname, extname, isAbsolute, join, relative, resolve, sep } from 'node:path'
import { atomicWriteFile } from './atomic-write'

export const ASSET_MANIFEST_FILENAME = '.genoffice-assets.json'

export interface OwnedAssetRecord {
  name: string
  sha256: string
}

export interface OwnedAssetManifest {
  version: 1
  files: OwnedAssetRecord[]
  /** Markdown file name -> app-owned files referenced by its last successful save. */
  documents: Record<string, string[]>
  /** Newly created files stay protected until their owning document successfully saves. */
  pending: Record<string, string[]>
}

export interface AssetRewrite {
  from: string
  to: string
}

export interface ReconcileResult {
  deleted: string[]
  preserved: string[]
  errors: string[]
}

export interface PreparedSaveAsAssets {
  targetDocumentPath: string
  text: string
  imageSources: string[]
  rewrites: AssetRewrite[]
  created: OwnedAssetRecord[]
}

const SHA256_RE = /^[a-f0-9]{64}$/
const manifestLocks = new Map<string, Promise<void>>()

function emptyManifest(): OwnedAssetManifest {
  return {
    version: 1,
    files: [],
    documents: Object.create(null) as Record<string, string[]>,
    pending: Object.create(null) as Record<string, string[]>,
  }
}

function assetsDirFor(documentPath: string): string {
  return join(dirname(resolve(documentPath)), 'assets')
}

function manifestPathFor(documentPath: string): string {
  return join(assetsDirFor(documentPath), ASSET_MANIFEST_FILENAME)
}

function documentKey(documentPath: string): string {
  return basename(resolve(documentPath))
}

function errorCode(error: unknown): string {
  return (error as NodeJS.ErrnoException | null)?.code ?? ''
}

function isSafeAssetName(name: string): boolean {
  return (
    name.length > 0 &&
    name.length <= 180 &&
    name !== '.' &&
    name !== '..' &&
    name !== ASSET_MANIFEST_FILENAME &&
    basename(name) === name &&
    !name.includes('/') &&
    !name.includes('\\') &&
    !name.includes('\0')
  )
}

function replaceControlCharacters(value: string): string {
  return [...value].map((character) => (character.charCodeAt(0) < 32 ? '_' : character)).join('')
}

function sanitizeAssetName(input: string): string {
  const raw = basename(input.replace(/\\/g, '/'))
  const rawExt = extname(raw)
  const ext = /^\.[a-z0-9]{1,10}$/i.test(rawExt) ? rawExt.toLowerCase() : ''
  const maxStem = 140 - ext.length
  const stem =
    replaceControlCharacters(raw.slice(0, raw.length - rawExt.length))
      .replace(/[/\\:*?"<>|#%]+/g, '_')
      .replace(/\s+/g, '_')
      .replace(/^\.+/, '')
      .slice(0, maxStem)
      .replace(/[. ]+$/g, '') || 'image'
  const name = `${stem}${ext}`
  return name === ASSET_MANIFEST_FILENAME ? `image-${name}` : name
}

function nameWithSuffix(name: string, suffix: number): string {
  if (suffix === 0) return name
  const ext = extname(name)
  return `${basename(name, ext)}-${suffix}${ext}`
}

function normalizedManifest(value: unknown): OwnedAssetManifest {
  if (!value || typeof value !== 'object') throw new Error('asset manifest is not an object')
  const raw = value as {
    version?: unknown
    files?: unknown
    documents?: unknown
    pending?: unknown
  }
  if (raw.version !== 1 || !Array.isArray(raw.files)) {
    throw new Error('asset manifest has an unsupported format')
  }

  const files: OwnedAssetRecord[] = []
  const seen = new Set<string>()
  for (const item of raw.files) {
    if (!item || typeof item !== 'object') throw new Error('asset manifest has a bad file entry')
    const { name, sha256 } = item as { name?: unknown; sha256?: unknown }
    if (
      typeof name !== 'string' ||
      !isSafeAssetName(name) ||
      typeof sha256 !== 'string' ||
      !SHA256_RE.test(sha256) ||
      seen.has(name)
    ) {
      throw new Error('asset manifest has an unsafe file entry')
    }
    seen.add(name)
    files.push({ name, sha256 })
  }

  const documents = Object.create(null) as Record<string, string[]>
  if (raw.documents !== undefined) {
    if (!raw.documents || typeof raw.documents !== 'object' || Array.isArray(raw.documents)) {
      throw new Error('asset manifest has bad document references')
    }
    for (const [key, names] of Object.entries(raw.documents as Record<string, unknown>)) {
      if (!isSafeAssetName(key) || !Array.isArray(names)) {
        throw new Error('asset manifest has unsafe document references')
      }
      const unique = new Set<string>()
      for (const name of names) {
        if (typeof name !== 'string' || !isSafeAssetName(name)) {
          throw new Error('asset manifest has unsafe document references')
        }
        unique.add(name)
      }
      documents[key] = [...unique].sort()
    }
  }

  const pending = Object.create(null) as Record<string, string[]>
  if (raw.pending !== undefined) {
    if (!raw.pending || typeof raw.pending !== 'object' || Array.isArray(raw.pending)) {
      throw new Error('asset manifest has bad pending references')
    }
    for (const [key, names] of Object.entries(raw.pending as Record<string, unknown>)) {
      if (!isSafeAssetName(key) || !Array.isArray(names)) {
        throw new Error('asset manifest has unsafe pending references')
      }
      const unique = new Set<string>()
      for (const name of names) {
        if (typeof name !== 'string' || !isSafeAssetName(name)) {
          throw new Error('asset manifest has unsafe pending references')
        }
        unique.add(name)
      }
      pending[key] = [...unique].sort()
    }
  }

  return {
    version: 1,
    files: files.sort((a, b) => a.name.localeCompare(b.name)),
    documents,
    pending,
  }
}

async function existingAssetsDir(documentPath: string): Promise<string | null> {
  const assetsDir = assetsDirFor(documentPath)
  try {
    const info = await lstat(assetsDir)
    if (!info.isDirectory() || info.isSymbolicLink()) {
      throw new Error('markdown assets path must be a real directory')
    }
    return assetsDir
  } catch (error) {
    if (errorCode(error) === 'ENOENT') return null
    throw error
  }
}

async function ensureAssetsDir(documentPath: string): Promise<string> {
  const existing = await existingAssetsDir(documentPath)
  if (existing) return existing
  const assetsDir = assetsDirFor(documentPath)
  await mkdir(assetsDir, { recursive: true })
  const info = await lstat(assetsDir)
  if (!info.isDirectory() || info.isSymbolicLink()) {
    throw new Error('markdown assets path must be a real directory')
  }
  return assetsDir
}

async function assertSafeManifestFile(path: string): Promise<void> {
  try {
    const info = await lstat(path)
    if (!info.isFile() || info.isSymbolicLink()) {
      throw new Error('asset manifest path must be a real file')
    }
  } catch (error) {
    if (errorCode(error) !== 'ENOENT') throw error
  }
}

async function readManifestUnlocked(documentPath: string): Promise<OwnedAssetManifest> {
  const assetsDir = await existingAssetsDir(documentPath)
  if (!assetsDir) return emptyManifest()
  const manifestPath = join(assetsDir, ASSET_MANIFEST_FILENAME)
  try {
    await assertSafeManifestFile(manifestPath)
    return normalizedManifest(JSON.parse(await readFile(manifestPath, 'utf8')))
  } catch (error) {
    if (errorCode(error) === 'ENOENT') return emptyManifest()
    if (error instanceof SyntaxError) {
      throw new Error('asset manifest is invalid JSON', { cause: error })
    }
    throw error
  }
}

async function writeManifestUnlocked(
  documentPath: string,
  manifest: OwnedAssetManifest,
): Promise<void> {
  const assetsDir = await ensureAssetsDir(documentPath)
  const normalized = normalizedManifest(manifest)
  const text = `${JSON.stringify(normalized, null, 2)}\n`
  const manifestPath = join(assetsDir, ASSET_MANIFEST_FILENAME)
  await assertSafeManifestFile(manifestPath)
  await atomicWriteFile(manifestPath, Buffer.from(text, 'utf8'))
}

async function withManifestLock<T>(documentPath: string, run: () => Promise<T>): Promise<T> {
  const key = manifestPathFor(documentPath)
  const previous = manifestLocks.get(key) ?? Promise.resolve()
  let release!: () => void
  const gate = new Promise<void>((resolveGate) => {
    release = resolveGate
  })
  const queued = previous.catch(() => {}).then(() => gate)
  manifestLocks.set(key, queued)
  await previous.catch(() => {})
  try {
    return await run()
  } finally {
    release()
    if (manifestLocks.get(key) === queued) manifestLocks.delete(key)
  }
}

async function sha256File(path: string): Promise<string> {
  return createHash('sha256')
    .update(await readFile(path))
    .digest('hex')
}

async function addOwnedRecord(documentPath: string, record: OwnedAssetRecord): Promise<void> {
  await withManifestLock(documentPath, async () => {
    const manifest = await readManifestUnlocked(documentPath)
    const existing = manifest.files.find((file) => file.name === record.name)
    if (existing && existing.sha256 !== record.sha256) {
      throw new Error('asset manifest already owns a different file with this name')
    }
    if (!existing) manifest.files.push(record)
    const key = documentKey(documentPath)
    manifest.pending[key] = [...new Set([...(manifest.pending[key] ?? []), record.name])].sort()
    await writeManifestUnlocked(documentPath, manifest)
  })
}

async function createOwnedAsset(
  documentPath: string,
  preferredName: string,
  create: (target: string) => Promise<void>,
): Promise<string> {
  const assetsDir = await ensureAssetsDir(documentPath)
  const base = sanitizeAssetName(preferredName)
  for (let suffix = 0; suffix < 100_000; suffix++) {
    const name = nameWithSuffix(base, suffix)
    const target = join(assetsDir, name)
    try {
      await create(target)
    } catch (error) {
      if (errorCode(error) === 'EEXIST') continue
      throw error
    }
    try {
      await addOwnedRecord(documentPath, { name, sha256: await sha256File(target) })
      return `assets/${name}`
    } catch (error) {
      await rm(target, { force: true }).catch(() => {})
      throw error
    }
  }
  throw new Error('markdown asset name space exhausted')
}

export async function copyImageIntoOwnedAssets(
  documentPath: string,
  sourcePath: string,
): Promise<string> {
  return createOwnedAsset(documentPath, basename(sourcePath), (target) =>
    copyFile(sourcePath, target, constants.COPYFILE_EXCL),
  )
}

export async function writeImageIntoOwnedAssets(
  documentPath: string,
  preferredName: string,
  bytes: Buffer,
): Promise<string> {
  return createOwnedAsset(documentPath, preferredName, (target) =>
    writeFile(target, bytes, { flag: 'wx' }),
  )
}

function decodeLocalSource(source: string): string | null {
  const trimmed = source.trim()
  if (
    !trimmed ||
    trimmed.includes('\0') ||
    trimmed.includes('?') ||
    trimmed.includes('#') ||
    trimmed.startsWith('/') ||
    trimmed.startsWith('\\') ||
    /^[a-zA-Z]:[\\/]/.test(trimmed) ||
    /^[a-z][a-z0-9+.-]*:/i.test(trimmed)
  ) {
    return null
  }
  try {
    return decodeURIComponent(trimmed)
  } catch {
    return null
  }
}

function isPathInside(root: string, candidate: string): boolean {
  const fromRoot = relative(root, candidate)
  return (
    fromRoot.length > 0 &&
    fromRoot !== '..' &&
    !fromRoot.startsWith(`..${sep}`) &&
    !isAbsolute(fromRoot)
  )
}

export async function resolveSafeRelativeImagePath(
  documentPath: string,
  source: string,
): Promise<string | null> {
  const decoded = decodeLocalSource(source)
  if (!decoded) return null
  const documentDir = dirname(resolve(documentPath))
  const candidate = resolve(documentDir, decoded.replace(/[\\/]/g, sep))
  if (!isPathInside(documentDir, candidate)) return null
  try {
    const [realDocumentDir, realCandidate, info] = await Promise.all([
      realpath(documentDir),
      realpath(candidate),
      stat(candidate),
    ])
    if (!info.isFile() || !isPathInside(realDocumentDir, realCandidate)) return null
    return realCandidate
  } catch {
    return null
  }
}

function ownedNameForSource(source: string): string | null {
  const decoded = decodeLocalSource(source)
  if (!decoded) return null
  const parts = decoded
    .replace(/\\/g, '/')
    .split('/')
    .filter((part) => part && part !== '.')
  if (
    parts.length !== 2 ||
    parts[0]!.toLowerCase() !== 'assets' ||
    parts.some((part) => part === '..') ||
    !isSafeAssetName(parts[1]!)
  ) {
    return null
  }
  return parts[1]!
}

async function siblingMarkdownReferences(
  documentPath: string,
  ownedNames: ReadonlySet<string>,
  includeCurrent = false,
): Promise<Set<string>> {
  const references = new Set<string>()
  const documentDir = dirname(resolve(documentPath))
  const currentKey = documentKey(documentPath)
  try {
    for (const entry of await readdir(documentDir, { withFileTypes: true })) {
      if (
        !entry.isFile() ||
        (!includeCurrent && entry.name === currentKey) ||
        !/\.(?:md|markdown)$/i.test(entry.name)
      ) {
        continue
      }
      const path = join(documentDir, entry.name)
      try {
        const info = await stat(path)
        // Refuse to load an unexpectedly huge sibling just for GC. Preserving
        // all candidates is the safe fallback when its references are unknown.
        if (info.size > 16 * 1024 * 1024) return new Set(ownedNames)
        const scan = scanImageSources(await readFile(path, 'utf8'))
        // A browser may still recover an image source from malformed HTML in a
        // way this deliberately small parser cannot prove. Keep every owned
        // candidate rather than collecting a possibly referenced file.
        if (scan.ambiguousHtml) return new Set(ownedNames)
        for (const { source } of scan.ranges) {
          const name = ownedNameForSource(source)
          if (name && ownedNames.has(name)) references.add(name)
        }
      } catch {
        return new Set(ownedNames)
      }
    }
  } catch {
    return new Set(ownedNames)
  }
  return references
}

interface DestinationRange {
  start: number
  end: number
  source: string
  htmlQuote?: '"' | "'" | null
}

interface ImageSourceScan {
  ranges: DestinationRange[]
  ambiguousHtml: boolean
}

function escapedAt(text: string, index: number): boolean {
  let slashes = 0
  for (let i = index - 1; i >= 0 && text[i] === '\\'; i--) slashes++
  return slashes % 2 === 1
}

function imageDestinationRanges(markdown: string): DestinationRange[] {
  const ranges: DestinationRange[] = []
  for (let i = 0; i < markdown.length - 2; i++) {
    if (markdown[i] !== '!' || markdown[i + 1] !== '[' || escapedAt(markdown, i)) continue
    let bracketDepth = 1
    let altEnd = i + 2
    for (; altEnd < markdown.length; altEnd++) {
      if (escapedAt(markdown, altEnd)) continue
      if (markdown[altEnd] === '[') bracketDepth++
      else if (markdown[altEnd] === ']' && --bracketDepth === 0) break
    }
    if (bracketDepth !== 0) continue
    let open = altEnd + 1
    while (markdown[open] === ' ' || markdown[open] === '\t') open++
    if (markdown[open] !== '(') {
      i = altEnd
      continue
    }

    let close = open + 1
    let parenDepth = 1
    let quote = ''
    let angle = false
    for (; close < markdown.length; close++) {
      const char = markdown[close]!
      if (escapedAt(markdown, close)) continue
      if (quote) {
        if (char === quote) quote = ''
        continue
      }
      if (angle) {
        if (char === '>') angle = false
        continue
      }
      if (char === '"' || char === "'") quote = char
      else if (char === '<') angle = true
      else if (char === '(') parenDepth++
      else if (char === ')' && --parenDepth === 0) break
    }
    if (parenDepth !== 0) continue

    let start = open + 1
    let end = close
    while (start < end && /\s/.test(markdown[start]!)) start++
    while (end > start && /\s/.test(markdown[end - 1]!)) end--
    if (markdown[start] === '<') {
      const angleEnd = markdown.indexOf('>', start + 1)
      if (angleEnd > start && angleEnd <= end) {
        start++
        end = angleEnd
      }
    } else {
      const raw = markdown.slice(start, end)
      const title = /\s+(?:"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'|\((?:\\.|[^)\\])*\))\s*$/.exec(raw)
      if (title?.index !== undefined) end = start + title.index
      while (end > start && /\s/.test(markdown[end - 1]!)) end--
    }
    if (end > start) {
      const source = markdown.slice(start, end).replace(/\\([\\()[\]<> ])/g, '$1')
      ranges.push({ start, end, source })
    }
    i = close
  }
  return ranges
}

const BASIC_HTML_ENTITIES: Readonly<Record<string, string>> = {
  amp: '&',
  apos: "'",
  gt: '>',
  lt: '<',
  nbsp: '\u00a0',
  quot: '"',
}

function decodeHtmlImageSource(value: string): { source: string; ambiguous: boolean } {
  let ambiguous = false
  const source = value.replace(/&([^&;\s]+);/g, (entity, body: string) => {
    if (body.startsWith('#')) {
      const hexadecimal = body[1]?.toLowerCase() === 'x'
      const digits = body.slice(hexadecimal ? 2 : 1)
      if (
        digits.length === 0 ||
        !(hexadecimal ? /^[0-9a-f]+$/i.test(digits) : /^[0-9]+$/.test(digits))
      ) {
        ambiguous = true
        return entity
      }
      const codePoint = Number.parseInt(digits, hexadecimal ? 16 : 10)
      if (codePoint <= 0 || codePoint > 0x10ffff || (codePoint >= 0xd800 && codePoint <= 0xdfff)) {
        ambiguous = true
        return entity
      }
      return String.fromCodePoint(codePoint)
    }
    const decoded = BASIC_HTML_ENTITIES[body.toLowerCase()]
    if (decoded === undefined) {
      // The complete HTML named-entity table is intentionally not duplicated
      // here. Unknown entities make sibling-GC parsing uncertain and therefore
      // trigger its retain-all fallback.
      ambiguous = true
      return entity
    }
    return decoded
  })
  if (/&#(?:x[0-9a-f]*|[0-9]*)/i.test(source)) ambiguous = true
  return { source, ambiguous }
}

function parseHtmlImageTag(
  text: string,
  tagStart: number,
): { nextIndex: number; range?: DestinationRange; ambiguous: boolean } {
  let cursor = tagStart + 4
  let sourceRange: DestinationRange | undefined
  let ambiguous = false

  while (cursor < text.length) {
    while (cursor < text.length && /\s/.test(text[cursor]!)) cursor += 1
    if (cursor >= text.length) return { nextIndex: text.length, ambiguous: true }
    if (text[cursor] === '>') {
      return {
        nextIndex: cursor + 1,
        ...(ambiguous || sourceRange === undefined ? {} : { range: sourceRange }),
        ambiguous,
      }
    }
    if (text[cursor] === '/' && text[cursor + 1] === '>') {
      return {
        nextIndex: cursor + 2,
        ...(ambiguous || sourceRange === undefined ? {} : { range: sourceRange }),
        ambiguous,
      }
    }
    if (text[cursor] === '/') {
      ambiguous = true
      cursor += 1
      continue
    }

    const nameStart = cursor
    while (
      cursor < text.length &&
      !/\s/.test(text[cursor]!) &&
      text[cursor] !== '=' &&
      text[cursor] !== '/' &&
      text[cursor] !== '>'
    ) {
      cursor += 1
    }
    if (cursor === nameStart) {
      ambiguous = true
      cursor += 1
      continue
    }
    const attributeName = text.slice(nameStart, cursor).toLowerCase()
    while (cursor < text.length && /\s/.test(text[cursor]!)) cursor += 1
    if (text[cursor] !== '=') {
      if (attributeName === 'src') ambiguous = true
      continue
    }

    cursor += 1
    while (cursor < text.length && /\s/.test(text[cursor]!)) cursor += 1
    if (
      cursor >= text.length ||
      text[cursor] === '>' ||
      (text[cursor] === '/' && text[cursor + 1] === '>')
    ) {
      if (attributeName === 'src') ambiguous = true
      continue
    }

    const quote: '"' | "'" | null = text[cursor] === '"' ? '"' : text[cursor] === "'" ? "'" : null
    let valueStart: number
    let valueEnd: number
    if (quote !== null) {
      valueStart = cursor + 1
      valueEnd = text.indexOf(quote, valueStart)
      if (valueEnd < 0) return { nextIndex: text.length, ambiguous: true }
      cursor = valueEnd + 1
    } else {
      valueStart = cursor
      while (
        cursor < text.length &&
        !/\s/.test(text[cursor]!) &&
        text[cursor] !== '>' &&
        !(text[cursor] === '/' && text[cursor + 1] === '>')
      ) {
        if (/["'`<=]/.test(text[cursor]!)) ambiguous = true
        cursor += 1
      }
      valueEnd = cursor
      if (valueEnd === valueStart && attributeName === 'src') ambiguous = true
    }

    if (attributeName !== 'src') continue
    if (sourceRange !== undefined) {
      ambiguous = true
      continue
    }
    const decoded = decodeHtmlImageSource(text.slice(valueStart, valueEnd))
    if (decoded.ambiguous) ambiguous = true
    sourceRange = {
      start: valueStart,
      end: valueEnd,
      source: decoded.source,
      htmlQuote: quote,
    }
  }
  return { nextIndex: text.length, ambiguous: true }
}

interface TextRange {
  start: number
  end: number
}

function positionInRanges(ranges: readonly TextRange[], position: number): TextRange | undefined {
  let low = 0
  let high = ranges.length - 1
  while (low <= high) {
    const middle = Math.floor((low + high) / 2)
    const range = ranges[middle]!
    if (position < range.start) high = middle - 1
    else if (position >= range.end) low = middle + 1
    else return range
  }
  return undefined
}

/** Fenced, indented and inline code ranges where image-looking text is not rendered content. */
function markdownCodeRanges(markdown: string): TextRange[] {
  const blockRanges: TextRange[] = []
  let fence: { marker: '`' | '~'; length: number; start: number } | null = null
  let offset = 0
  for (const lineWithBreak of markdown.match(/.*(?:\n|$)/g) ?? []) {
    if (!lineWithBreak) continue
    const line = lineWithBreak.replace(/\r?\n$/, '')
    const lineEnd = offset + lineWithBreak.length
    if (fence) {
      const close = new RegExp(`^ {0,3}\\${fence.marker}{${fence.length},}[ \\t]*$`)
      if (close.test(line)) {
        blockRanges.push({ start: fence.start, end: lineEnd })
        fence = null
      }
    } else {
      const open = /^ {0,3}(`{3,}|~{3,})(.*)$/.exec(line)
      if (open && (open[1]![0] === '~' || !open[2]!.includes('`'))) {
        fence = {
          marker: open[1]![0] as '`' | '~',
          length: open[1]!.length,
          start: offset,
        }
      } else if (/^(?: {4}|\t)/.test(line)) {
        blockRanges.push({ start: offset, end: lineEnd })
      }
    }
    offset = lineEnd
  }
  if (fence) blockRanges.push({ start: fence.start, end: markdown.length })
  blockRanges.sort((left, right) => left.start - right.start)

  const ranges = [...blockRanges]
  for (let index = 0; index < markdown.length;) {
    const block = positionInRanges(blockRanges, index)
    if (block) {
      index = block.end
      continue
    }
    if (markdown[index] !== '`' || escapedAt(markdown, index)) {
      index += 1
      continue
    }
    let runEnd = index + 1
    while (markdown[runEnd] === '`') runEnd += 1
    const runLength = runEnd - index
    let search = runEnd
    let closeEnd = -1
    while (search < markdown.length) {
      const blocked = positionInRanges(blockRanges, search)
      if (blocked) {
        search = blocked.end
        continue
      }
      const next = markdown.indexOf('`', search)
      if (next < 0) break
      let nextEnd = next + 1
      while (markdown[nextEnd] === '`') nextEnd += 1
      if (nextEnd - next === runLength) {
        closeEnd = nextEnd
        break
      }
      search = nextEnd
    }
    if (closeEnd > 0) {
      ranges.push({ start: index, end: closeEnd })
      index = closeEnd
    } else {
      // An unmatched run is literal Markdown, so later HTML may still render.
      index = runEnd
    }
  }
  return ranges.sort((left, right) => left.start - right.start)
}

function htmlImageSourceRanges(
  markdown: string,
  codeRanges: readonly TextRange[],
): ImageSourceScan {
  const ranges: DestinationRange[] = []
  let ambiguousHtml = false
  for (let index = 0; index < markdown.length; index += 1) {
    const code = positionInRanges(codeRanges, index)
    if (code) {
      index = code.end - 1
      continue
    }
    if (markdown[index] !== '<') continue
    if (markdown.startsWith('<!--', index)) {
      const commentEnd = markdown.indexOf('-->', index + 4)
      if (commentEnd < 0) {
        if (/<img(?:\s|\/|>)/i.test(markdown.slice(index + 4))) ambiguousHtml = true
        break
      }
      index = commentEnd + 2
      continue
    }
    if (
      markdown.slice(index + 1, index + 4).toLowerCase() !== 'img' ||
      (index + 4 < markdown.length && !/[\s/>]/.test(markdown[index + 4]!))
    ) {
      continue
    }
    const parsed = parseHtmlImageTag(markdown, index)
    if (parsed.ambiguous) ambiguousHtml = true
    if (parsed.range) ranges.push(parsed.range)
    index = Math.max(index, parsed.nextIndex - 1)
  }
  return { ranges, ambiguousHtml }
}

function scanImageSources(markdown: string): ImageSourceScan {
  const codeRanges = markdownCodeRanges(markdown)
  const html = htmlImageSourceRanges(markdown, codeRanges)
  const markdownRanges = imageDestinationRanges(markdown).filter(
    (range) => positionInRanges(codeRanges, range.start) === undefined,
  )
  const ordered = [...markdownRanges, ...html.ranges].sort(
    (left, right) => left.start - right.start || left.end - right.end,
  )
  const ranges: DestinationRange[] = []
  for (const range of ordered) {
    const previous = ranges[ranges.length - 1]
    if (previous && range.start < previous.end) {
      html.ambiguousHtml = true
      continue
    }
    ranges.push(range)
  }
  return { ranges, ambiguousHtml: html.ambiguousHtml }
}

export function extractMarkdownImageSources(markdown: string): string[] {
  return scanImageSources(markdown).ranges.map((range) => range.source)
}

function encodeHtmlAttributeReplacement(value: string, quote: '"' | "'" | null): string {
  let encoded = ''
  for (const character of value) {
    if (character === '&') encoded += '&amp;'
    else if (character === '<') encoded += '&lt;'
    else if (quote === '"' && character === '"') encoded += '&quot;'
    else if (quote === "'" && character === "'") encoded += '&#39;'
    else if (quote === null && /[\s"'`=>]/.test(character)) {
      encoded += `&#${character.codePointAt(0)!};`
    } else encoded += character
  }
  return encoded
}

export function rewriteMarkdownImageSources(
  markdown: string,
  rewrites: ReadonlyMap<string, string>,
): string {
  if (rewrites.size === 0) return markdown
  const ranges = scanImageSources(markdown).ranges
  let cursor = 0
  let output = ''
  for (const range of ranges) {
    const replacement = rewrites.get(range.source)
    if (replacement === undefined) continue
    output += markdown.slice(cursor, range.start)
    output +=
      range.htmlQuote === undefined
        ? replacement
        : encodeHtmlAttributeReplacement(replacement, range.htmlQuote)
    cursor = range.end
  }
  return cursor === 0 ? markdown : output + markdown.slice(cursor)
}

async function removeCreatedRecords(
  documentPath: string,
  created: readonly OwnedAssetRecord[],
): Promise<void> {
  if (created.length === 0) return
  await withManifestLock(documentPath, async () => {
    const manifest = await readManifestUnlocked(documentPath)
    const createdNames = new Set(created.map((record) => record.name))
    const key = documentKey(documentPath)
    manifest.pending[key] = (manifest.pending[key] ?? []).filter((name) => !createdNames.has(name))
    if (manifest.pending[key]!.length === 0) delete manifest.pending[key]
    const referenced = new Set([
      ...Object.values(manifest.documents).flat(),
      ...Object.values(manifest.pending).flat(),
    ])
    const removeNames = new Set<string>()
    for (const record of created) {
      if (referenced.has(record.name)) continue
      const current = manifest.files.find(
        (file) => file.name === record.name && file.sha256 === record.sha256,
      )
      if (!current) continue
      const path = join(assetsDirFor(documentPath), record.name)
      try {
        const info = await lstat(path)
        if (!info.isFile() || info.isSymbolicLink() || (await sha256File(path)) !== record.sha256) {
          continue
        }
        await rm(path)
        removeNames.add(record.name)
      } catch (error) {
        if (errorCode(error) === 'ENOENT') removeNames.add(record.name)
      }
    }
    manifest.files = manifest.files.filter((file) => !removeNames.has(file.name))
    for (const key of Object.keys(manifest.documents)) {
      manifest.documents[key] = manifest.documents[key]!.filter((name) => !removeNames.has(name))
      if (manifest.documents[key]!.length === 0) delete manifest.documents[key]
    }
    for (const key of Object.keys(manifest.pending)) {
      manifest.pending[key] = manifest.pending[key]!.filter((name) => !removeNames.has(name))
      if (manifest.pending[key]!.length === 0) delete manifest.pending[key]
    }
    await writeManifestUnlocked(documentPath, manifest)
  })
}

export async function prepareAssetsForSaveAs(
  sourceDocumentPath: string,
  targetDocumentPath: string,
  text: string,
  imageSources: readonly string[],
): Promise<PreparedSaveAsAssets> {
  const allImageSources = [...imageSources]
  const knownSources = new Set(allImageSources)
  for (const source of extractMarkdownImageSources(text)) {
    if (knownSources.has(source)) continue
    knownSources.add(source)
    allImageSources.push(source)
  }
  if (resolve(dirname(sourceDocumentPath)) === resolve(dirname(targetDocumentPath))) {
    return {
      targetDocumentPath,
      text,
      imageSources: allImageSources,
      rewrites: [],
      created: [],
    }
  }

  const created: OwnedAssetRecord[] = []
  const rewrites = new Map<string, string>()
  const targetAssetsDir = await ensureAssetsDir(targetDocumentPath)
  try {
    await withManifestLock(targetDocumentPath, async () => {
      const manifest = await readManifestUnlocked(targetDocumentPath)
      const reservedNames = new Set(manifest.files.map((file) => file.name))
      const copiedByRealPath = new Map<string, string>()
      for (const source of allImageSources) {
        if (rewrites.has(source)) continue
        const safeSource = await resolveSafeRelativeImagePath(sourceDocumentPath, source)
        if (!safeSource) continue
        const alreadyCopied = copiedByRealPath.get(safeSource)
        if (alreadyCopied) {
          rewrites.set(source, alreadyCopied)
          continue
        }

        const base = sanitizeAssetName(basename(safeSource))
        let copied: OwnedAssetRecord | null = null
        for (let suffix = 0; suffix < 100_000; suffix++) {
          const name = nameWithSuffix(base, suffix)
          if (reservedNames.has(name)) continue
          const target = join(targetAssetsDir, name)
          try {
            await copyFile(safeSource, target, constants.COPYFILE_EXCL)
            copied = { name, sha256: await sha256File(target) }
            break
          } catch (error) {
            if (errorCode(error) === 'EEXIST') continue
            throw error
          }
        }
        if (!copied) throw new Error('markdown asset name space exhausted')
        created.push(copied)
        manifest.files.push(copied)
        reservedNames.add(copied.name)
        const targetKey = documentKey(targetDocumentPath)
        manifest.pending[targetKey] = [
          ...new Set([...(manifest.pending[targetKey] ?? []), copied.name]),
        ].sort()
        const authored = `assets/${copied.name}`
        copiedByRealPath.set(safeSource, authored)
        rewrites.set(source, authored)
      }
      const rewrittenDestinations = new Set(
        extractMarkdownImageSources(rewriteMarkdownImageSources(text, rewrites)),
      )
      for (const destination of rewrites.values()) {
        if (!rewrittenDestinations.has(destination)) {
          throw new Error('could not safely rewrite a Markdown image during Save As')
        }
      }
      await writeManifestUnlocked(targetDocumentPath, manifest)
    })
  } catch (error) {
    for (const record of created) {
      await rm(join(targetAssetsDir, record.name), { force: true }).catch(() => {})
    }
    throw error
  }

  const rewriteList = [...rewrites].map(([from, to]) => ({ from, to }))
  return {
    targetDocumentPath,
    text: rewriteMarkdownImageSources(text, rewrites),
    imageSources: allImageSources.map((source) => rewrites.get(source) ?? source),
    rewrites: rewriteList,
    created,
  }
}

export async function rollbackPreparedSaveAsAssets(plan: PreparedSaveAsAssets): Promise<void> {
  await removeCreatedRecords(plan.targetDocumentPath, plan.created)
}

export async function pendingOwnedAssetsForDocument(documentPath: string): Promise<string[]> {
  try {
    return await withManifestLock(documentPath, async () => {
      const manifest = await readManifestUnlocked(documentPath)
      return [...(manifest.pending[documentKey(documentPath)] ?? [])]
    })
  } catch {
    return []
  }
}

async function resolvePendingOwnedAssets(
  documentPath: string,
  pendingAtOperationStart?: readonly string[],
): Promise<ReconcileResult> {
  const result: ReconcileResult = { deleted: [], preserved: [], errors: [] }
  try {
    return await withManifestLock(documentPath, async () => {
      const manifest = await readManifestUnlocked(documentPath)
      const key = documentKey(documentPath)
      const currentPending = manifest.pending[key] ?? []
      const requested = new Set(pendingAtOperationStart ?? currentPending)
      const pendingNames = new Set(currentPending.filter((name) => requested.has(name)))
      if (pendingNames.size === 0) return result
      manifest.pending[key] = currentPending.filter((name) => !pendingNames.has(name))
      if (manifest.pending[key]!.length === 0) delete manifest.pending[key]

      const assetsDir = await existingAssetsDir(documentPath)
      if (!assetsDir) return result
      const recordsByName = new Map(manifest.files.map((record) => [record.name, record]))
      const ownedNames = new Set(recordsByName.keys())
      const referenced = new Set([
        ...Object.values(manifest.documents).flat(),
        ...Object.values(manifest.pending).flat(),
      ])
      for (const name of await siblingMarkdownReferences(documentPath, ownedNames, true)) {
        referenced.add(name)
      }

      const relinquished = new Set<string>()
      for (const name of pendingNames) {
        const record = recordsByName.get(name)
        if (!record) continue
        const path = join(assetsDir, name)
        try {
          const info = await lstat(path)
          if (
            !info.isFile() ||
            info.isSymbolicLink() ||
            (await sha256File(path)) !== record.sha256
          ) {
            relinquished.add(name)
            result.preserved.push(name)
            continue
          }
          if (referenced.has(name)) {
            result.preserved.push(name)
            continue
          }
          await rm(path)
          relinquished.add(name)
          result.deleted.push(name)
        } catch (error) {
          if (errorCode(error) === 'ENOENT') {
            relinquished.add(name)
          } else {
            result.preserved.push(name)
            result.errors.push(`${name}: ${error instanceof Error ? error.message : String(error)}`)
          }
        }
      }

      manifest.files = manifest.files.filter((record) => !relinquished.has(record.name))
      for (const references of [manifest.documents, manifest.pending]) {
        for (const referenceKey of Object.keys(references)) {
          references[referenceKey] = references[referenceKey]!.filter(
            (name) => !relinquished.has(name),
          )
          if (references[referenceKey]!.length === 0) delete references[referenceKey]
        }
      }
      await writeManifestUnlocked(documentPath, manifest)
      return result
    })
  } catch (error) {
    result.errors.push(error instanceof Error ? error.message : String(error))
    return result
  }
}

/**
 * The user explicitly discarded unsaved edits. Remove only this document's
 * pending app-owned files; saved/shared references and changed files survive.
 */
export async function discardPendingOwnedAssets(documentPath: string): Promise<ReconcileResult> {
  return resolvePendingOwnedAssets(documentPath)
}

/**
 * Save As moved the editing session to a new path. Resolve only pending files
 * that existed when Save As started; later insertions remain undoable.
 */
export async function resolveSourcePendingAfterSaveAs(
  sourceDocumentPath: string,
  pendingAtSaveStart: readonly string[],
): Promise<ReconcileResult> {
  return resolvePendingOwnedAssets(sourceDocumentPath, pendingAtSaveStart)
}

export async function reconcileOwnedAssets(
  documentPath: string,
  imageSources: readonly string[],
  options: { pendingNames?: readonly string[] } = {},
): Promise<ReconcileResult> {
  const result: ReconcileResult = { deleted: [], preserved: [], errors: [] }
  try {
    return await withManifestLock(documentPath, async () => {
      const manifest = await readManifestUnlocked(documentPath)
      const assetsDir = await existingAssetsDir(documentPath)
      if (!assetsDir) return result

      const validRecords: OwnedAssetRecord[] = []
      const relinquished = new Set<string>()
      for (const record of manifest.files) {
        const path = join(assetsDir, record.name)
        try {
          const info = await lstat(path)
          if (
            !info.isFile() ||
            info.isSymbolicLink() ||
            (await sha256File(path)) !== record.sha256
          ) {
            relinquished.add(record.name)
            result.preserved.push(record.name)
          } else {
            validRecords.push(record)
          }
        } catch (error) {
          if (errorCode(error) !== 'ENOENT') {
            relinquished.add(record.name)
            result.preserved.push(record.name)
            result.errors.push(
              `${record.name}: ${error instanceof Error ? error.message : String(error)}`,
            )
          }
        }
      }
      manifest.files = validRecords

      const ownedNames = new Set(validRecords.map((record) => record.name))
      for (const key of Object.keys(manifest.documents)) {
        manifest.documents[key] = manifest.documents[key]!.filter(
          (name) => ownedNames.has(name) && !relinquished.has(name),
        )
        if (manifest.documents[key]!.length === 0) delete manifest.documents[key]
      }
      for (const key of Object.keys(manifest.pending)) {
        manifest.pending[key] = manifest.pending[key]!.filter(
          (name) => ownedNames.has(name) && !relinquished.has(name),
        )
        if (manifest.pending[key]!.length === 0) delete manifest.pending[key]
      }

      const currentReferences = new Set<string>()
      for (const source of imageSources) {
        const name = ownedNameForSource(source)
        if (name && ownedNames.has(name)) currentReferences.add(name)
      }
      const key = documentKey(documentPath)
      if (currentReferences.size > 0) {
        manifest.documents[key] = [...currentReferences].sort()
      } else {
        delete manifest.documents[key]
      }
      const pendingToResolve = new Set(options.pendingNames ?? manifest.pending[key] ?? [])
      manifest.pending[key] = (manifest.pending[key] ?? []).filter(
        (name) => !pendingToResolve.has(name),
      )
      if (manifest.pending[key]!.length === 0) delete manifest.pending[key]

      const referenced = new Set([
        ...Object.values(manifest.documents).flat(),
        ...Object.values(manifest.pending).flat(),
      ])
      for (const name of await siblingMarkdownReferences(documentPath, ownedNames)) {
        referenced.add(name)
      }
      const deleted = new Set<string>()
      for (const record of validRecords) {
        if (referenced.has(record.name)) continue
        const path = join(assetsDir, record.name)
        try {
          const info = await lstat(path)
          if (
            !info.isFile() ||
            info.isSymbolicLink() ||
            (await sha256File(path)) !== record.sha256
          ) {
            result.preserved.push(record.name)
            continue
          }
          await rm(path)
          deleted.add(record.name)
          result.deleted.push(record.name)
        } catch (error) {
          if (errorCode(error) === 'ENOENT') {
            deleted.add(record.name)
          } else {
            result.errors.push(
              `${record.name}: ${error instanceof Error ? error.message : String(error)}`,
            )
          }
        }
      }
      manifest.files = manifest.files.filter((record) => !deleted.has(record.name))
      for (const doc of Object.keys(manifest.documents)) {
        manifest.documents[doc] = manifest.documents[doc]!.filter((name) => !deleted.has(name))
        if (manifest.documents[doc]!.length === 0) delete manifest.documents[doc]
      }
      for (const doc of Object.keys(manifest.pending)) {
        manifest.pending[doc] = manifest.pending[doc]!.filter((name) => !deleted.has(name))
        if (manifest.pending[doc]!.length === 0) delete manifest.pending[doc]
      }
      try {
        await writeManifestUnlocked(documentPath, manifest)
      } catch (error) {
        result.errors.push(error instanceof Error ? error.message : String(error))
      }
      return result
    })
  } catch (error) {
    result.errors.push(error instanceof Error ? error.message : String(error))
    return result
  }
}

export async function renameOwnedAssetDocument(
  oldDocumentPath: string,
  newDocumentPath: string,
): Promise<void> {
  if (resolve(dirname(oldDocumentPath)) !== resolve(dirname(newDocumentPath))) return
  await withManifestLock(newDocumentPath, async () => {
    const manifest = await readManifestUnlocked(newDocumentPath)
    const oldKey = documentKey(oldDocumentPath)
    const newKey = documentKey(newDocumentPath)
    const oldReferences = manifest.documents[oldKey]
    const oldPending = manifest.pending[oldKey]
    if ((!oldReferences && !oldPending) || oldKey === newKey) return
    if (oldReferences) {
      manifest.documents[newKey] = [
        ...new Set([...(manifest.documents[newKey] ?? []), ...oldReferences]),
      ].sort()
      delete manifest.documents[oldKey]
    }
    if (oldPending) {
      manifest.pending[newKey] = [
        ...new Set([...(manifest.pending[newKey] ?? []), ...oldPending]),
      ].sort()
      delete manifest.pending[oldKey]
    }
    await writeManifestUnlocked(newDocumentPath, manifest)
  })
}

export async function readOwnedAssetManifest(
  documentPath: string,
): Promise<OwnedAssetManifest | null> {
  try {
    const assetsDir = await existingAssetsDir(documentPath)
    if (!assetsDir) return null
    return await readManifestUnlocked(documentPath)
  } catch {
    return null
  }
}

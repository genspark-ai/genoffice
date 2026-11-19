import { randomUUID } from 'node:crypto'
import { mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises'
import { basename, extname, isAbsolute, join, relative, resolve, sep } from 'node:path'

export const MARKDOWN_CONVERSION_TTL_MS = 7 * 24 * 60 * 60 * 1000
const SESSION_MARKER = '.genoffice-markdown-conversion.json'

interface ConversionSessionMarker {
  version: 1
  createdAt: number
}

function errorCode(error: unknown): string {
  return (error as NodeJS.ErrnoException | null)?.code ?? ''
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

function replaceControlCharacters(value: string): string {
  return [...value].map((character) => (character.charCodeAt(0) < 32 ? '_' : character)).join('')
}

function safeBaseName(input: string): string {
  const raw = basename(input)
  const ext = extname(raw)
  return (
    replaceControlCharacters(raw.slice(0, raw.length - ext.length))
      .replace(/[/\\:*?"<>|]+/g, '_')
      .replace(/\s+/g, '_')
      .replace(/^\.+/, '')
      .slice(0, 100)
      .replace(/[. ]+$/g, '') || 'Untitled'
  )
}

async function readSessionMarker(directory: string): Promise<ConversionSessionMarker | null> {
  try {
    const value = JSON.parse(await readFile(join(directory, SESSION_MARKER), 'utf8')) as {
      version?: unknown
      createdAt?: unknown
    }
    if (
      value.version !== 1 ||
      typeof value.createdAt !== 'number' ||
      !Number.isFinite(value.createdAt) ||
      value.createdAt < 0
    ) {
      return null
    }
    return { version: 1, createdAt: value.createdAt }
  } catch {
    return null
  }
}

/**
 * Removes only marked GenOffice session directories older than the TTL.
 * Unknown files/directories under the cache root are deliberately ignored.
 */
export async function cleanupStaleMarkdownConversions(
  root: string,
  now = Date.now(),
  ttlMs = MARKDOWN_CONVERSION_TTL_MS,
): Promise<string[]> {
  const absoluteRoot = resolve(root)
  await mkdir(absoluteRoot, { recursive: true })
  const removed: string[] = []
  for (const entry of await readdir(absoluteRoot, { withFileTypes: true })) {
    if (!entry.isDirectory() || !entry.name.startsWith('session-')) continue
    const directory = join(absoluteRoot, entry.name)
    if (!isPathInside(absoluteRoot, directory)) continue
    const marker = await readSessionMarker(directory)
    if (!marker || now - marker.createdAt < ttlMs) continue
    await rm(directory, { recursive: true, force: true })
    removed.push(directory)
  }
  return removed
}

export async function createMarkdownConversionSession(
  root: string,
  options: { now?: number; id?: string; ttlMs?: number } = {},
): Promise<string> {
  const now = options.now ?? Date.now()
  const absoluteRoot = resolve(root)
  await cleanupStaleMarkdownConversions(
    absoluteRoot,
    now,
    options.ttlMs ?? MARKDOWN_CONVERSION_TTL_MS,
  )
  const rawId = options.id ?? randomUUID()
  const id = rawId.replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 80) || randomUUID()
  const directory = join(absoluteRoot, `session-${now}-${id}`)
  if (!isPathInside(absoluteRoot, directory)) {
    throw new Error('invalid Markdown conversion session path')
  }
  await mkdir(directory)
  const marker: ConversionSessionMarker = { version: 1, createdAt: now }
  await writeFile(join(directory, SESSION_MARKER), `${JSON.stringify(marker)}\n`, { flag: 'wx' })
  return directory
}

export async function writeMarkdownConversion(
  sessionDirectory: string,
  suggestedName: string,
  bytes: Buffer,
  id = randomUUID(),
): Promise<string> {
  const marker = await readSessionMarker(sessionDirectory)
  if (!marker) throw new Error('invalid Markdown conversion session')
  const base = safeBaseName(suggestedName)
  const safeId = id.replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 80) || randomUUID()
  for (let suffix = 0; suffix < 100_000; suffix++) {
    const tail = suffix === 0 ? '' : `-${suffix}`
    const target = join(sessionDirectory, `${base}-${safeId}${tail}.docx`)
    if (!isPathInside(resolve(sessionDirectory), resolve(target))) {
      throw new Error('invalid Markdown conversion output path')
    }
    try {
      await writeFile(target, bytes, { flag: 'wx' })
      return target
    } catch (error) {
      if (errorCode(error) === 'EEXIST') continue
      throw error
    }
  }
  throw new Error('Markdown conversion name space exhausted')
}

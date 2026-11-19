import { mkdtemp, mkdir, readdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join, relative } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import {
  MARKDOWN_CONVERSION_TTL_MS,
  cleanupStaleMarkdownConversions,
  createMarkdownConversionSession,
  writeMarkdownConversion,
} from '../src/main/conversion-lifecycle'

const tempDirectories: string[] = []

async function temporaryDirectory(prefix: string): Promise<string> {
  const directory = await mkdtemp(join(tmpdir(), prefix))
  tempDirectories.push(directory)
  return directory
}

afterEach(async () => {
  await Promise.all(
    tempDirectories.splice(0).map((directory) => rm(directory, { recursive: true, force: true })),
  )
})

describe('Markdown convert-and-open cache', () => {
  it('puts repeated conversions in an app-owned session instead of beside the Markdown file', async () => {
    const sourceDirectory = await temporaryDirectory('markdown-convert-source-')
    const cacheRoot = await temporaryDirectory('markdown-convert-cache-')
    await writeFile(join(sourceDirectory, 'notes.md'), '# Notes')
    const session = await createMarkdownConversionSession(cacheRoot, {
      now: 1_000,
      id: 'test-session',
    })

    const first = await writeMarkdownConversion(
      session,
      'notes',
      Buffer.from('first docx'),
      'first',
    )
    const second = await writeMarkdownConversion(
      session,
      'notes',
      Buffer.from('second docx'),
      'second',
    )

    expect(first).not.toBe(second)
    expect(dirname(first)).toBe(session)
    expect(dirname(second)).toBe(session)
    expect(await readdir(sourceDirectory)).toEqual(['notes.md'])
  })

  it('cleans only marked sessions after the explicit TTL', async () => {
    const cacheRoot = await temporaryDirectory('markdown-convert-ttl-')
    const old = await createMarkdownConversionSession(cacheRoot, { now: 1, id: 'old' })
    const unknown = join(cacheRoot, 'session-user-data')
    await mkdir(unknown)
    await writeFile(join(unknown, 'keep.txt'), 'keep')

    const removed = await cleanupStaleMarkdownConversions(cacheRoot, MARKDOWN_CONVERSION_TTL_MS + 2)

    expect(removed).toEqual([old])
    expect(await readdir(unknown)).toEqual(['keep.txt'])
  })

  it('sanitizes generated names so output cannot escape the marked session', async () => {
    const cacheRoot = await temporaryDirectory('markdown-convert-safe-')
    const session = await createMarkdownConversionSession(cacheRoot, { now: 1, id: 'safe' })
    const output = await writeMarkdownConversion(
      session,
      '../../outside',
      Buffer.from('docx'),
      '../bad-id',
    )

    expect(relative(session, output).startsWith('..')).toBe(false)
    expect(output.endsWith('.docx')).toBe(true)
  })
})

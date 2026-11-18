// @vitest-environment node
import { constants, existsSync, rmSync, writeFileSync } from 'node:fs'
import { copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { copyImageIntoOwnedAssets, prepareAssetsForSaveAs } from '../src/main/asset-lifecycle'

vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return { ...actual, copyFile: vi.fn(actual.copyFile) }
})

const actualCopyFile = (
  await vi.importActual<typeof import('node:fs/promises')>('node:fs/promises')
).copyFile

const tempDirectories: string[] = []

async function temporaryDirectory(prefix: string): Promise<string> {
  const directory = await mkdtemp(join(tmpdir(), prefix))
  tempDirectories.push(directory)
  return directory
}

afterEach(async () => {
  vi.mocked(copyFile).mockImplementation(actualCopyFile)
  await Promise.all(
    tempDirectories.splice(0).map((directory) => rm(directory, { recursive: true, force: true })),
  )
})

/**
 * A mount that refuses the platform copy, the way libuv fails on it: an
 * exclusive copy still refuses an existing file, any other copy has created the
 * destination by the time it fails and removes it on the way out.
 */
function platformCopyFailsLikeLibuv(): void {
  vi.mocked(copyFile).mockImplementation(async (_source, target, mode) => {
    const path = String(target)
    if (((mode ?? 0) & constants.COPYFILE_EXCL) !== 0 && existsSync(path))
      throw Object.assign(new Error('EEXIST: injected'), { code: 'EEXIST' })
    writeFileSync(path, '')
    rmSync(path)
    throw Object.assign(new Error('ENOTSUP: injected'), { code: 'ENOTSUP' })
  })
}

describe.skipIf(process.platform === 'win32')(
  'HTML image copies where the platform copy does not work',
  () => {
    it('lands a picked image beside the document with its bytes', async () => {
      const directory = await temporaryDirectory('html-assets-copy-')
      const source = join(directory, 'picked.png')
      await writeFile(source, 'picked bytes')
      platformCopyFailsLikeLibuv()

      const reference = await copyImageIntoOwnedAssets(join(directory, 'doc.html'), source)

      expect(reference).toBe('assets/picked.png')
      expect(await readFile(join(directory, reference), 'utf8')).toBe('picked bytes')
    })

    it('moves on to the next name on a name collision and keeps the existing file', async () => {
      const directory = await temporaryDirectory('html-assets-copy-')
      const source = join(directory, 'picked.png')
      await writeFile(source, 'picked bytes')
      await mkdir(join(directory, 'assets'))
      await writeFile(join(directory, 'assets', 'picked.png'), 'already here')
      platformCopyFailsLikeLibuv()

      const reference = await copyImageIntoOwnedAssets(join(directory, 'doc.html'), source)

      expect(reference).toBe('assets/picked-1.png')
      expect(await readFile(join(directory, reference), 'utf8')).toBe('picked bytes')
      expect(await readFile(join(directory, 'assets', 'picked.png'), 'utf8')).toBe('already here')
    })

    it('copies Save As images and keeps a different file that holds the preferred name', async () => {
      const sourceDirectory = await temporaryDirectory('html-assets-source-')
      const targetDirectory = await temporaryDirectory('html-assets-target-')
      await mkdir(join(sourceDirectory, 'images'))
      await writeFile(join(sourceDirectory, 'images', 'manual.png'), 'manual source')
      await mkdir(join(targetDirectory, 'assets'))
      await writeFile(join(targetDirectory, 'assets', 'manual.png'), 'target user file')
      platformCopyFailsLikeLibuv()

      const prepared = await prepareAssetsForSaveAs(
        join(sourceDirectory, 'source.html'),
        join(targetDirectory, 'copy.html'),
        '<img src="images/manual.png">',
        ['images/manual.png'],
      )

      expect(prepared.rewrites).toEqual([{ from: 'images/manual.png', to: 'assets/manual-1.png' }])
      expect(await readFile(join(targetDirectory, 'assets', 'manual-1.png'), 'utf8')).toBe(
        'manual source',
      )
      expect(await readFile(join(sourceDirectory, 'images', 'manual.png'), 'utf8')).toBe(
        'manual source',
      )
      expect(await readFile(join(targetDirectory, 'assets', 'manual.png'), 'utf8')).toBe(
        'target user file',
      )
    })
  },
)

import { mkdtemp, readdir, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it } from 'vitest'
import { ImageExportSessions } from '../src/main/image-export'

const roots: string[] = []
afterEach(async () => {
  await Promise.all(roots.map((root) => rm(root, { recursive: true, force: true })))
  roots.length = 0
})
const png =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLttAAAAABJRU5ErkJggg=='
async function setup() {
  const root = await mkdtemp(join(tmpdir(), 'markdown-images-test-'))
  roots.push(root)
  const sessions = new ImageExportSessions()
  const id = await sessions.start(1, root, 'Report')
  return { root, sessions, id }
}

it('writes numbered pages into a new folder without overwriting another export', async () => {
  const { root, sessions, id } = await setup()
  await sessions.write(1, id, 1, png)
  await sessions.write(1, id, 2, png)
  const dir = await sessions.finish(1, id, true)
  expect(await readdir(dir!)).toEqual(['page-01.png', 'page-02.png'])
  expect(await readFile(join(dir!, 'page-01.png'))).toEqual(Buffer.from(png, 'base64'))
  await sessions.start(1, root, 'Report')
  expect((await readdir(root)).length).toBe(2)
  await sessions.dispose(1)
  expect(await readdir(root)).toEqual([dir!.split(/[\\/]/).pop()])
})

it('rejects writes from another renderer, invalid pages and invalid PNG data', async () => {
  const { sessions, id } = await setup()
  await expect(sessions.write(2, id, 1, png)).rejects.toThrow(/authorized/)
  await expect(sessions.write(1, id, -1, png)).rejects.toThrow(/page/)
  await expect(sessions.write(1, id, 1, 'not png')).rejects.toThrow(/PNG/)
  await sessions.write(1, id, 1, png)
  await expect(sessions.write(1, id, 1, png)).rejects.toThrow(/page/)
})

it('removes partial images when export fails and revokes the session', async () => {
  const { root, sessions, id } = await setup()
  await sessions.write(1, id, 1, png)
  expect(await sessions.finish(1, id, false)).toBeNull()
  expect(await readdir(root)).toEqual([])
  await expect(sessions.write(1, id, 2, png)).rejects.toThrow(/authorized/)
})

it('disposes only the closing renderer’s unfinished export', async () => {
  const { root, sessions, id } = await setup()
  const other = await sessions.start(2, root, '../other')
  await sessions.dispose(1)
  await expect(sessions.write(1, id, 1, png)).rejects.toThrow(/authorized/)
  await sessions.write(2, other, 1, png)
  const dir = await sessions.finish(2, other, true)
  expect(dir!.startsWith(root + '/')).toBe(true)
  expect(await readdir(dir!)).toEqual(['page-01.png'])
})

it('does not report an empty export as successful', async () => {
  const { root, sessions, id } = await setup()
  await expect(sessions.finish(1, id, true)).rejects.toThrow(/empty/)
  expect(await readdir(root)).toEqual([])
})

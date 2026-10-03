import { mkdtempSync, mkdirSync, rmSync, symlinkSync, writeFileSync } from 'node:fs'
import { homedir, tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, describe, expect, it } from 'vitest'
import { isLocalMediaPathAllowed, loadMediaReference, localMediaRoots } from '../src/media-tools'

// A local media path reaches the loader straight from the model's tool call in
// the desktop apps, so the extension check alone decides what leaves the
// machine. The default roots are the working directory and the temp directory;
// `outside` is deliberately under neither (the temp dir is the only writable
// place a test may use that is guaranteed to sit outside the working tree).
const outside = mkdtempSync(join(homedir(), '.genoffice-aisearch-outside-'))
const inside = mkdtempSync(join(tmpdir(), 'genoffice-aisearch-inside-'))
const secret = join(outside, 'private.png')

afterAll(() => {
  rmSync(outside, { recursive: true, force: true })
  rmSync(inside, { recursive: true, force: true })
})

function writePng(path: string): string {
  writeFileSync(path, Buffer.alloc(16, 7))
  return path
}

describe('loadMediaReference local paths', () => {
  it('reads a media file inside an allowed root', async () => {
    const path = writePng(join(inside, 'shot.png'))
    const blob = await loadMediaReference(path)
    expect(blob.mime).toBe('image/png')
    expect(blob.bytes.byteLength).toBe(16)
  })

  it('refuses a media file outside every allowed root', async () => {
    const path = writePng(secret)
    await expect(loadMediaReference(path)).rejects.toThrow(/allowed/i)
  })

  it('refuses a symlink inside an allowed root that points outside it', async () => {
    writePng(secret)
    const link = join(inside, 'looks-local.png')
    symlinkSync(secret, link)
    await expect(loadMediaReference(link)).rejects.toThrow(/allowed/i)
  })

  it('refuses a file reached through a symlinked directory', async () => {
    writePng(secret)
    const dirLink = join(inside, 'linked-dir')
    symlinkSync(outside, dirLink, 'dir')
    await expect(loadMediaReference(join(dirLink, 'private.png'))).rejects.toThrow(/allowed/i)
  })

  it('still refuses a non-media extension inside an allowed root', async () => {
    const path = join(inside, 'notes.env')
    writeFileSync(path, 'SECRET=1')
    await expect(loadMediaReference(path)).rejects.toThrow(/Unsupported media file/)
  })

  it('reports a missing file rather than a missing root', async () => {
    mkdirSync(join(inside, 'empty'), { recursive: true })
    await expect(loadMediaReference(join(inside, 'empty', 'gone.png'))).rejects.toThrow(
      /File not found/,
    )
  })

  it('honours a caller-supplied root list', async () => {
    const path = writePng(secret)
    // a caller that knows where the open document lives can widen the allowlist
    await expect(loadMediaReference(path, [outside])).resolves.toMatchObject({ mime: 'image/png' })
    await expect(loadMediaReference(path)).rejects.toThrow(/allowed/i)
  })

  it('matches whole path segments, not string prefixes', () => {
    const root = join(tmpdir(), 'genoffice-aisearch-root')
    const sibling = `${root}-evil`
    mkdirSync(root, { recursive: true })
    mkdirSync(sibling, { recursive: true })
    try {
      const insideRoot = writePng(join(root, 'ok.png'))
      const besideRoot = writePng(join(sibling, 'sneaky.png'))
      expect(isLocalMediaPathAllowed(insideRoot, [root])).toBe(true)
      expect(isLocalMediaPathAllowed(besideRoot, [root])).toBe(false)
    } finally {
      rmSync(root, { recursive: true, force: true })
      rmSync(sibling, { recursive: true, force: true })
    }
  })

  it('defaults to the working directory and the temp directory', () => {
    expect(localMediaRoots()).toEqual([process.cwd(), tmpdir()])
  })
})

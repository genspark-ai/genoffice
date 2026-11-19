import { mkdtempSync, mkdirSync, rmSync, symlinkSync, writeFileSync } from 'node:fs'
import { homedir, tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, describe, expect, it } from 'vitest'
import {
  documentMediaRoots,
  isLocalMediaPathAllowed,
  loadMediaReference,
  localMediaRoots,
} from '../src/media-tools'

// A local media path reaches the loader straight from the model's tool call in
// the desktop apps, so what leaves the machine is decided here. The allowlist
// is opt-in: with no roots the extension check alone applies (the pre-allowlist
// behaviour), and a caller that knows where the open document lives passes its
// own roots. `outside` is deliberately under neither the working tree nor the
// temp dir (homedir is the one writable place a test may use that is
// guaranteed to sit outside both), so it exercises the no-roots path.
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

describe('localMediaRoots', () => {
  it('has no default: an allowlist nobody opts into relocates the trust boundary', () => {
    // The bug this replaced defaulted to [process.cwd(), tmpdir()], which
    // refused legitimate use from any other directory and was a no-op when a
    // packaged macOS app's cwd is "/". There is no default now.
    expect(localMediaRoots()).toEqual([])
  })

  it('drops empty entries so a caller can pass an optional directory', () => {
    expect(localMediaRoots(undefined, '/keep', null, '')).toEqual(['/keep'])
  })
})

describe('documentMediaRoots', () => {
  it("is the open document's directory plus the attachment directory", () => {
    expect(documentMediaRoots(join('/work', 'deck.pptx'), '/attach')).toEqual(['/work', '/attach'])
  })

  it('keeps only the attachment directory for an untitled document', () => {
    expect(documentMediaRoots(undefined, '/attach')).toEqual(['/attach'])
    expect(documentMediaRoots(null, undefined)).toEqual([])
  })
})

describe('loadMediaReference with no roots (extension-only)', () => {
  it('reads a media file anywhere, including outside cwd and tmpdir', async () => {
    // No allowlist configured: the extension check is the only gate, which is
    // exactly the pre-fix behaviour. This is the regression guard for the
    // removed [cwd, tmpdir] default.
    const path = writePng(secret)
    expect(outside.startsWith(process.cwd())).toBe(false)
    const blob = await loadMediaReference(path)
    expect(blob.mime).toBe('image/png')
    expect(blob.bytes.byteLength).toBe(16)
  })

  it('reports a readable path as allowed when no roots are configured', () => {
    expect(isLocalMediaPathAllowed(secret)).toBe(true)
  })

  it('still refuses a non-media extension', async () => {
    const path = join(outside, 'notes.env')
    writeFileSync(path, 'SECRET=1')
    await expect(loadMediaReference(path)).rejects.toThrow(/Unsupported media file/)
  })
})

describe('loadMediaReference with caller-supplied roots', () => {
  it('reads a media file inside a root', async () => {
    const path = writePng(join(inside, 'shot.png'))
    const blob = await loadMediaReference(path, [inside])
    expect(blob.mime).toBe('image/png')
    expect(blob.bytes.byteLength).toBe(16)
  })

  it('refuses a media file outside every root', async () => {
    const path = writePng(secret)
    await expect(loadMediaReference(path, [inside])).rejects.toThrow(/allowed/i)
  })

  it('refuses a symlink inside a root that points outside it', async () => {
    writePng(secret)
    const link = join(inside, 'looks-local.png')
    symlinkSync(secret, link)
    await expect(loadMediaReference(link, [inside])).rejects.toThrow(/allowed/i)
  })

  it('refuses a file reached through a symlinked directory', async () => {
    writePng(secret)
    const dirLink = join(inside, 'linked-dir')
    symlinkSync(outside, dirLink, 'dir')
    await expect(loadMediaReference(join(dirLink, 'private.png'), [inside])).rejects.toThrow(
      /allowed/i,
    )
  })

  it('still refuses a non-media extension inside a root', async () => {
    const path = join(inside, 'notes.env')
    writeFileSync(path, 'SECRET=1')
    await expect(loadMediaReference(path, [inside])).rejects.toThrow(/Unsupported media file/)
  })

  it('reports a missing file rather than a missing root', async () => {
    mkdirSync(join(inside, 'empty'), { recursive: true })
    await expect(loadMediaReference(join(inside, 'empty', 'gone.png'), [inside])).rejects.toThrow(
      /File not found/,
    )
  })

  it('honours a caller-supplied root list that names the file directory', async () => {
    const path = writePng(secret)
    // the CLI shape: the user named this file, so its own directory is the root
    await expect(loadMediaReference(path, [outside])).resolves.toMatchObject({ mime: 'image/png' })
    await expect(loadMediaReference(path, [inside])).rejects.toThrow(/allowed/i)
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
})

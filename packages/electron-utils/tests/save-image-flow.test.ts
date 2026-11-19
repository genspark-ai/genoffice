import { mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'

const picked: { filePath: string } = { filePath: '' }
const fetched = vi.fn()
vi.mock('electron', () => ({
  dialog: {
    // called with (options) when there is no parent window, (win, options) otherwise
    showSaveDialog: vi.fn(async (...args: unknown[]) => {
      const opts = args[args.length - 1] as { defaultPath?: string }
      return {
        canceled: false,
        filePath: picked.filePath || opts.defaultPath || 'x',
      }
    }),
  },
  net: { fetch: fetched },
}))

import { saveImageFromUrl } from '../src/save-image'

describe('saveImageFromUrl', () => {
  it('decodes a data URL and writes the bytes to the picked path', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'save-image-'))
    picked.filePath = join(dir, 'out.png')
    const res = await saveImageFromUrl(null, 'data:image/png;base64,aGVsbG8=', { title: 't' })
    expect(res).toEqual({ ok: true, path: picked.filePath })
    expect((await readFile(picked.filePath)).toString()).toBe('hello')
  })

  it('fetches non-data URLs through net.fetch and names the file from the URL', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'save-image-'))
    picked.filePath = ''
    fetched.mockResolvedValueOnce(
      new Response(new Uint8Array([1, 2, 3]), { headers: { 'content-type': 'image/png' } }),
    )
    const { dialog } = await import('electron')
    const res = await saveImageFromUrl(null, 'md-asset:///d/assets/pic.png', {
      title: 't',
      fallbackDir: dir,
    })
    expect(fetched).toHaveBeenCalledWith('md-asset:///d/assets/pic.png')
    expect(res).toEqual({ ok: true, path: expect.any(String) })
    expect(vi.mocked(dialog.showSaveDialog).mock.lastCall?.at(-1)).toMatchObject({
      defaultPath: expect.stringMatching(/pic\.png$/),
      filters: [{ name: 'PNG', extensions: ['png'] }],
    })
  })

  it('refuses file URLs without touching the dialog', async () => {
    const res = await saveImageFromUrl(null, 'file:///etc/hosts', { title: 't' })
    expect(res.ok).toBe(false)
    expect(fetched).toHaveBeenCalledTimes(1)
  })
})

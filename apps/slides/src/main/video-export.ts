import { open, unlink, type FileHandle } from 'node:fs/promises'
import type { IpcMain } from 'electron'
import type { VideoFileWrite } from '../shared/ipc'

const handles = new Map<number, { fh: FileHandle; path: string }>()
let nextId = 1

/** Streamed MP4 writes from the renderer's encoder; positions are absolute (the muxer patches headers at the end). */
export function registerVideoExportIpc(ipcMain: IpcMain): void {
  ipcMain.handle('slides:video-file-open', async (_e, path: string): Promise<number> => {
    const fh = await open(path, 'w')
    const id = nextId++
    handles.set(id, { fh, path })
    return id
  })
  ipcMain.handle('slides:video-file-write', async (_e, op: VideoFileWrite): Promise<void> => {
    const h = handles.get(op.id)
    if (!h) throw new Error('video export: file handle closed')
    const buf = Buffer.from(op.data.buffer, op.data.byteOffset, op.data.byteLength)
    let done = 0
    while (done < buf.length) {
      const { bytesWritten } = await h.fh.write(buf, done, buf.length - done, op.position + done)
      done += bytesWritten
    }
  })
  ipcMain.handle(
    'slides:video-file-close',
    async (_e, id: number, discard: boolean): Promise<void> => {
      const h = handles.get(id)
      if (!h) return
      handles.delete(id)
      await h.fh.close()
      if (discard) await unlink(h.path).catch(() => undefined)
    },
  )
}

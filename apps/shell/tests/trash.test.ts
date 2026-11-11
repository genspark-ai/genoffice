import { EventEmitter } from 'node:events'
import { beforeEach, describe, expect, it, vi } from 'vitest'

const mockSpawn = vi.fn()
const mockGetPath = vi.fn()
const mockOpenPath = vi.fn()

vi.mock('node:child_process', () => ({
  spawn: (...args: unknown[]) => mockSpawn(...args),
}))

vi.mock('electron', () => ({
  app: {
    getPath: (...args: unknown[]) => mockGetPath(...args),
  },
  shell: {
    openPath: (...args: unknown[]) => mockOpenPath(...args),
  },
}))

import { openTrash } from '../src/main/trash'

describe('openTrash', () => {
  beforeEach(() => {
    vi.clearAllMocks()
    mockGetPath.mockReturnValue('/Users/test')
  })

  describe('Windows (win32)', () => {
    it('opens RecycleBinFolder using explorer.exe with detached process', async () => {
      const fakeChild = new EventEmitter() as EventEmitter & { unref: ReturnType<typeof vi.fn> }
      fakeChild.unref = vi.fn()

      mockSpawn.mockImplementation(() => {
        queueMicrotask(() => fakeChild.emit('spawn'))
        return fakeChild
      })

      const result = await openTrash('win32')
      expect(result).toBe(true)
      expect(mockSpawn).toHaveBeenCalledWith('explorer.exe', ['shell:RecycleBinFolder'], {
        detached: true,
        stdio: 'ignore',
      })
      expect(fakeChild.unref).toHaveBeenCalled()
    })

    it('returns false if explorer.exe fails to spawn', async () => {
      const fakeChild = new EventEmitter() as EventEmitter & { unref: ReturnType<typeof vi.fn> }
      fakeChild.unref = vi.fn()

      mockSpawn.mockImplementation(() => {
        queueMicrotask(() => fakeChild.emit('error', new Error('spawn failed')))
        return fakeChild
      })

      const result = await openTrash('win32')
      expect(result).toBe(false)
    })
  })

  describe('macOS (darwin)', () => {
    it('opens ~/.Trash using shell.openPath', async () => {
      mockOpenPath.mockResolvedValue('')

      const result = await openTrash('darwin')
      expect(result).toBe(true)
      expect(mockGetPath).toHaveBeenCalledWith('home')
      expect(mockOpenPath).toHaveBeenCalledWith(expect.stringMatching(/[\\/]\.Trash$/))
    })

    it('returns false when shell.openPath returns an error string', async () => {
      mockOpenPath.mockResolvedValue('Failed to open directory')

      const result = await openTrash('darwin')
      expect(result).toBe(false)
    })
  })

  describe('Linux (linux)', () => {
    it('spawns xdg-open with trash:/// detached and returns true on spawn', async () => {
      const fakeChild = new EventEmitter() as EventEmitter & { unref: ReturnType<typeof vi.fn> }
      fakeChild.unref = vi.fn()

      mockSpawn.mockImplementation(() => {
        queueMicrotask(() => fakeChild.emit('spawn'))
        return fakeChild
      })

      const result = await openTrash('linux')
      expect(result).toBe(true)
      expect(mockSpawn).toHaveBeenCalledWith('xdg-open', ['trash:///'], {
        detached: true,
        stdio: 'ignore',
      })
      expect(fakeChild.unref).toHaveBeenCalled()
    })

    it('returns false when xdg-open fails to spawn', async () => {
      const fakeChild = new EventEmitter() as EventEmitter & { unref: ReturnType<typeof vi.fn> }
      fakeChild.unref = vi.fn()

      mockSpawn.mockImplementation(() => {
        queueMicrotask(() => fakeChild.emit('error', new Error('spawn failed')))
        return fakeChild
      })

      const result = await openTrash('linux')
      expect(result).toBe(false)
    })
  })
})

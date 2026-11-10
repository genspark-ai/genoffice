import { EventEmitter } from 'node:events'
import { beforeEach, describe, expect, it, vi } from 'vitest'

const mockSpawn = vi.fn()
const mockExecFile = vi.fn()
const mockOpenPath = vi.fn()

vi.mock('node:child_process', () => ({
  spawn: (...args: unknown[]) => mockSpawn(...args),
  execFile: (...args: unknown[]) => mockExecFile(...args),
}))

vi.mock('electron', () => ({
  shell: {
    openPath: (...args: unknown[]) => mockOpenPath(...args),
  },
}))

import { openTrash } from '../src/main/trash'

describe('openTrash', () => {
  beforeEach(() => {
    vi.clearAllMocks()
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
    it('uses AppleScript via osascript to ask Finder to open and activate trash', async () => {
      mockExecFile.mockImplementation((_cmd, _args, callback) => {
        callback(null, '', '')
      })

      const result = await openTrash('darwin')
      expect(result).toBe(true)
      expect(mockExecFile).toHaveBeenCalledWith(
        'osascript',
        [
          '-e',
          'tell application "Finder" to open trash',
          '-e',
          'tell application "Finder" to activate',
        ],
        expect.any(Function),
      )
    })

    it('returns false when osascript encounters an error', async () => {
      mockExecFile.mockImplementation((_cmd, _args, callback) => {
        callback(new Error('osascript failure'), '', '')
      })

      const result = await openTrash('darwin')
      expect(result).toBe(false)
    })
  })

  describe('Linux (linux)', () => {
    it('opens trash via xdg-open trash:/// when available', async () => {
      mockExecFile.mockImplementation((cmd, args, callback) => {
        if (cmd === 'xdg-open' && args[0] === 'trash:///') {
          callback(null, '', '')
        }
      })

      const result = await openTrash('linux')
      expect(result).toBe(true)
      expect(mockExecFile).toHaveBeenCalledWith('xdg-open', ['trash:///'], expect.any(Function))
      expect(mockOpenPath).not.toHaveBeenCalled()
    })

    it('falls back to shell.openPath with XDG_DATA_HOME/Trash/files if xdg-open fails', async () => {
      mockExecFile.mockImplementation((_cmd, _args, callback) => {
        callback(new Error('xdg-open failed'), '', '')
      })
      mockOpenPath.mockResolvedValue('')

      const result = await openTrash('linux')
      expect(result).toBe(true)
      expect(mockOpenPath).toHaveBeenCalledWith(expect.stringContaining('Trash'))
    })
  })
})

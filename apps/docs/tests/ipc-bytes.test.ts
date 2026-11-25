import { describe, expect, it } from 'vitest'
import { ipcBytes } from '../src/main/ipc-bytes'

describe('bytes sent over IPC', () => {
  it('are read from an ArrayBuffer and from any view of one', () => {
    const jpeg = [0xff, 0xd8, 0xff, 0xe0]
    expect([...ipcBytes(Uint8Array.from(jpeg).buffer)!]).toEqual(jpeg)
    expect([...ipcBytes(Buffer.from(jpeg))!]).toEqual(jpeg)
    // a view onto part of a larger buffer is only that part
    const padded = Uint8Array.from([0, 0, ...jpeg, 0])
    expect([...ipcBytes(padded.subarray(2, 6))!]).toEqual(jpeg)
    expect([...ipcBytes(new DataView(padded.buffer, 2, 4))!]).toEqual(jpeg)
  })

  it('are nothing for anything else', () => {
    expect(ipcBytes(undefined)).toBeNull()
    expect(ipcBytes('ffd8')).toBeNull()
    expect(ipcBytes([0xff, 0xd8])).toBeNull()
  })
})

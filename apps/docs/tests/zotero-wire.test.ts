import { describe, expect, it } from 'vitest'
import { encodeZoteroFrame, ZoteroFrameDecoder } from '../src/main/zotero-wire'

describe('Zotero integration wire framing', () => {
  it('encodes the transaction ID and UTF-8 byte length as big-endian uint32 values', () => {
    const frame = encodeZoteroFrame(42, '中文')
    expect(frame.readUInt32BE(0)).toBe(42)
    expect(frame.readUInt32BE(4)).toBe(6)
    expect(frame.subarray(8).toString('utf8')).toBe('中文')
  })

  it('decodes a frame split across arbitrary TCP chunks', () => {
    const frame = encodeZoteroFrame(7, '["Document_complete",[1]]')
    const decoder = new ZoteroFrameDecoder()
    expect(decoder.push(frame.subarray(0, 3))).toEqual([])
    expect(decoder.push(frame.subarray(3, 11))).toEqual([])
    expect(decoder.push(frame.subarray(11))).toEqual([
      { transactionId: 7, payload: '["Document_complete",[1]]' },
    ])
  })

  it('decodes multiple frames delivered in one TCP chunk', () => {
    const decoder = new ZoteroFrameDecoder()
    const chunk = Buffer.concat([encodeZoteroFrame(1, 'one'), encodeZoteroFrame(2, 'two')])
    expect(decoder.push(chunk)).toEqual([
      { transactionId: 1, payload: 'one' },
      { transactionId: 2, payload: 'two' },
    ])
  })
})

import { createServer, type AddressInfo, type Server } from 'node:net'
import { describe, expect, it } from 'vitest'
import { encodeZoteroFrame, ZoteroFrameDecoder, ZoteroWireClient } from '../src/main/zotero-wire'

describe('Zotero integration wire framing', () => {
  it('encodes the transaction ID and UTF-8 byte length as big-endian uint32 values', () => {
    const cjk = String.fromCharCode(0x4e2d, 0x6587)
    const frame = encodeZoteroFrame(42, cjk)
    expect(frame.readUInt32BE(0)).toBe(42)
    expect(frame.readUInt32BE(4)).toBe(6)
    expect(frame.subarray(8).toString('utf8')).toBe(cjk)
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

describe('ZoteroWireClient', () => {
  const listen = (server: Server) =>
    new Promise<number>((resolve) =>
      server.listen(0, '127.0.0.1', () => resolve((server.address() as AddressInfo).port)),
    )
  const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms))

  it('services document requests and settles only when Zotero closes the connection', async () => {
    const seen: string[] = []
    const server = createServer((socket) => {
      const decoder = new ZoteroFrameDecoder()
      socket.on('data', (chunk) => {
        for (const frame of decoder.push(chunk)) {
          seen.push(frame.payload)
          if (frame.transactionId === 0) {
            setTimeout(
              () => socket.write(encodeZoteroFrame(1, '["Document_getDocumentData",[1]]')),
              30,
            )
          } else if (frame.transactionId === 1) {
            socket.write(encodeZoteroFrame(2, '["Document_complete",[1]]'))
          }
        }
      })
      socket.on('end', () => socket.end())
    })
    const port = await listen(server)
    try {
      const client = new ZoteroWireClient('127.0.0.1', port)
      let settled = false
      const run = client
        .run('refresh', async (request) =>
          request.command === 'Document_getDocumentData' ? '<data/>' : null,
        )
        .then(() => {
          settled = true
        })
      await sleep(10)
      expect(settled).toBe(false)
      expect(client.isActive).toBe(true)
      await run
      expect(client.isActive).toBe(false)
      expect(seen[0]).toContain('"command":"refresh"')
      expect(seen).toContain('"<data/>"')
      expect(seen).toContain('null')
    } finally {
      await new Promise((resolve) => server.close(resolve))
    }
  })

  it('rejects when nothing listens on the Zotero port', async () => {
    const server = createServer()
    const port = await listen(server)
    await new Promise((resolve) => server.close(resolve))
    const client = new ZoteroWireClient('127.0.0.1', port)
    await expect(client.run('refresh', async () => null)).rejects.toThrow(/ECONNREFUSED/)
    expect(client.isActive).toBe(false)
  })
})

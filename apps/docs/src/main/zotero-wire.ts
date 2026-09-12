import { createConnection, type Socket } from 'node:net'

export const ZOTERO_INTEGRATION_HOST = '127.0.0.1'
export const ZOTERO_INTEGRATION_PORT = 23116
const MAX_PAYLOAD_BYTES = 16 * 1024 * 1024

export interface ZoteroWireFrame {
  transactionId: number
  payload: string
}

export interface ZoteroWireRequest {
  command: string
  args: unknown[]
}

export function encodeZoteroFrame(transactionId: number, payload: string): Buffer {
  const body = Buffer.from(payload, 'utf8')
  const frame = Buffer.allocUnsafe(8 + body.length)
  frame.writeUInt32BE(transactionId >>> 0, 0)
  frame.writeUInt32BE(body.length, 4)
  body.copy(frame, 8)
  return frame
}

export class ZoteroFrameDecoder {
  private buffered = Buffer.alloc(0)

  push(chunk: Uint8Array): ZoteroWireFrame[] {
    this.buffered = Buffer.concat([this.buffered, Buffer.from(chunk)])
    const frames: ZoteroWireFrame[] = []
    while (this.buffered.length >= 8) {
      const transactionId = this.buffered.readUInt32BE(0)
      const length = this.buffered.readUInt32BE(4)
      if (length > MAX_PAYLOAD_BYTES) throw new Error('Zotero integration payload is too large')
      if (this.buffered.length < 8 + length) break
      frames.push({
        transactionId,
        payload: this.buffered.subarray(8, 8 + length).toString('utf8'),
      })
      this.buffered = this.buffered.subarray(8 + length)
    }
    return frames
  }
}

function parseRequest(payload: string): ZoteroWireRequest {
  const value: unknown = JSON.parse(payload)
  if (!Array.isArray(value) || typeof value[0] !== 'string' || !Array.isArray(value[1])) {
    throw new Error('Invalid Zotero integration request')
  }
  return { command: value[0], args: value[1] }
}

function errorPayload(error: unknown): string {
  return `ERR:${error instanceof Error ? error.message : String(error)}`
}

export class ZoteroWireClient {
  private socket: Socket | null = null
  private active = false

  get isActive(): boolean {
    return this.active
  }

  async run(
    command: string,
    handleRequest: (request: ZoteroWireRequest) => Promise<unknown>,
  ): Promise<void> {
    if (this.active) throw new Error('A Zotero operation is already running')
    this.active = true
    const decoder = new ZoteroFrameDecoder()

    try {
      const socket = await new Promise<Socket>((resolve, reject) => {
        const candidate = createConnection({
          host: ZOTERO_INTEGRATION_HOST,
          port: ZOTERO_INTEGRATION_PORT,
        })
        const onError = (error: Error) => reject(error)
        candidate.once('error', onError)
        candidate.once('connect', () => {
          candidate.off('error', onError)
          resolve(candidate)
        })
      })
      this.socket = socket
      socket.on('data', (chunk) => {
        let frames: ZoteroWireFrame[]
        try {
          frames = decoder.push(chunk)
        } catch (error) {
          socket.destroy(error instanceof Error ? error : new Error(String(error)))
          return
        }
        for (const frame of frames) void this.respondToFrame(socket, frame, handleRequest)
      })
      socket.once('close', () => {
        if (this.socket === socket) this.socket = null
        this.active = false
      })
      socket.once('error', () => {
        // Post-connect errors terminate the socket; request failures are returned as ERR frames.
      })
      socket.write(encodeZoteroFrame(0, JSON.stringify({ command, templateVersion: 1 })))
    } catch (error) {
      this.active = false
      this.socket = null
      throw error
    }
  }

  close(): void {
    this.socket?.destroy()
    this.socket = null
    this.active = false
  }

  private async respondToFrame(
    socket: Socket,
    frame: ZoteroWireFrame,
    handleRequest: (request: ZoteroWireRequest) => Promise<unknown>,
  ): Promise<void> {
    if (frame.transactionId === 0) return
    let payload: string
    let command = ''
    try {
      const request = parseRequest(frame.payload)
      command = request.command
      payload = JSON.stringify(await handleRequest(request))
    } catch (error) {
      payload = errorPayload(error)
    }
    if (!socket.destroyed) socket.write(encodeZoteroFrame(frame.transactionId, payload))
    if (command === 'Document_complete') socket.end()
  }
}

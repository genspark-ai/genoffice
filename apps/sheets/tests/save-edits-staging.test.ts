/**
 * Renderer-side staging of a save's cell edits: small sets ride inline,
 * larger sets are uploaded in bounded ordered chunks and replaced by a
 * transfer id in the save request.
 */
import { describe, expect, it } from 'vitest'
import { stageEditsForSave, type SaveEditsTransferApi } from '../src/renderer/save-edits-staging'
import type { WorkbookCellEdit } from '../src/shared/desktop-api'
import {
  MAX_SAVE_EDITS,
  MAX_SAVE_EDITS_TOTAL,
  SAVE_EDITS_CHUNK_MAX,
} from '../src/shared/ipc-channels'

const SESSION = '5d4f6f7a-1c2b-4e3d-9a8f-0b1c2d3e4f5a'

function edits(count: number): WorkbookCellEdit[] {
  return Array.from({ length: count }, (_, index) => ({
    sheetId: 'sheet-1',
    row: index % 1_048_576,
    column: 0,
    writeValue: true,
    value: index,
  }))
}

interface RecordedChunk {
  transferId: string
  seq: number
  count: number
}

function recordingApi(options: { failChunkAt?: number } = {}) {
  const begins: { transferId: string; total: number }[] = []
  const chunks: RecordedChunk[] = []
  const aborts: string[] = []
  const api: SaveEditsTransferApi = {
    async beginSaveEditsTransfer(request) {
      begins.push({ transferId: request.transferId, total: request.total })
    },
    async sendSaveEditsChunk(request) {
      if (request.seq === options.failChunkAt) throw new Error('chunk upload failed')
      chunks.push({ transferId: request.transferId, seq: request.seq, count: request.edits.length })
    },
    async abortSaveEditsTransfer(request) {
      aborts.push(request.transferId)
    },
  }
  return { api, begins, chunks, aborts }
}

describe('stageEditsForSave', () => {
  it('keeps small edit sets inline without opening a transfer', async () => {
    const { api, begins, chunks } = recordingApi()
    const inline = edits(10)
    const staged = await stageEditsForSave(api, SESSION, inline)
    expect(staged.edits).toBe(inline)
    expect(staged.editsTransferId).toBeUndefined()
    expect(begins).toHaveLength(0)
    expect(chunks).toHaveLength(0)
  })

  it('chunks edit sets above the inline cap', async () => {
    const { api, begins, chunks } = recordingApi()
    const total = MAX_SAVE_EDITS + SAVE_EDITS_CHUNK_MAX + 1
    const staged = await stageEditsForSave(api, SESSION, edits(total))
    expect(staged.edits).toEqual([])
    expect(staged.editsTransferId).toBe(begins[0]?.transferId)
    expect(begins[0]?.total).toBe(total)
    // Ordered slices covering exactly the input, each within the chunk cap.
    expect(chunks.map((chunk) => chunk.seq)).toEqual(chunks.map((_, index) => index))
    expect(chunks.every((chunk) => chunk.count <= SAVE_EDITS_CHUNK_MAX)).toBe(true)
    expect(chunks.reduce((sum, chunk) => sum + chunk.count, 0)).toBe(total)
    expect(chunks.at(-1)?.count).toBe(1)
  })

  it('aborts the transfer when a chunk upload fails', async () => {
    const { api, begins, aborts } = recordingApi({ failChunkAt: 1 })
    await expect(
      stageEditsForSave(api, SESSION, edits(MAX_SAVE_EDITS + SAVE_EDITS_CHUNK_MAX + 1)),
    ).rejects.toThrow(/chunk upload failed/)
    expect(aborts).toEqual([begins[0]?.transferId])
  })

  it('rejects edit sets above the absolute ceiling', async () => {
    const { api, begins } = recordingApi()
    // Length-lie stub: building 10M real objects would slow the suite down.
    const oversized = { length: MAX_SAVE_EDITS_TOTAL + 1 } as unknown as WorkbookCellEdit[]
    await expect(stageEditsForSave(api, SESSION, oversized)).rejects.toThrow(/maximum/)
    expect(begins).toHaveLength(0)
  })
})

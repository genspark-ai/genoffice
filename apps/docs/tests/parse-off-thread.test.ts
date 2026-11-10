import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { parseDocx, readSections, saveDocx } from '@genoffice/docx-engine'
import { parseDocxOffThread } from '../src/renderer/parse-off-thread'
import { blocksToPmDoc, pmDocOptions, pmDocToSavePlan } from '../src/renderer/editor/convert'

const fixture = () =>
  new Uint8Array(readFileSync(join(__dirname, 'pagination-corpus/docx/kitchen-sink.docx')))

describe('parse off the UI thread', () => {
  it('falls back to the inline parser where Worker is unavailable', async () => {
    expect(typeof Worker).toBe('undefined')
    const parsed = await parseDocxOffThread(fixture())
    expect(parsed.blocks.length).toBeGreaterThan(10)
  })

  it('the worker handoff (structured clone + transferred buffers) keeps the model usable', async () => {
    const bytes = fixture()
    const byteLength = bytes.byteLength
    const parsed = await parseDocx(bytes)
    // what the worker posts back: a clone with the source bytes and font buffers moved
    const clone = structuredClone(parsed, {
      transfer: [
        parsed.internal.originalBytes.buffer as ArrayBuffer,
        ...(parsed.embeddedFonts ?? []).map((f) => f.data.buffer as ArrayBuffer),
      ],
    })
    // the source bytes travelled with the clone (the caller's view is detached)
    expect(clone.internal.originalBytes.byteLength).toBe(byteLength)
    expect(bytes.byteLength).toBe(0)
    expect(clone.styles instanceof Map).toBe(true)
    const pm = blocksToPmDoc(clone.blocks, readSections(clone), pmDocOptions(clone))
    expect(pm.content?.length ?? 0).toBeGreaterThan(10)
    const plan = pmDocToSavePlan(pm, clone.blocks)
    const saved = await saveDocx(clone, plan.saveBlocks, {})
    // jsdom hands structuredClone results back from another realm; a real Worker
    // message lands in the page's own — re-wrap so JSZip's type check sees ours
    const reparsed = await parseDocx(new Uint8Array(saved))
    expect(reparsed.blocks.length).toBe(parsed.blocks.length)
    expect(reparsed.blocks.map((b) => b.docxIndex)).toEqual(parsed.blocks.map((b) => b.docxIndex))
  })
})

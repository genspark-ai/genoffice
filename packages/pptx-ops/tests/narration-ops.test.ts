/** addNarration / clearNarration: one transaction each, never offered to the AI. */
import { describe, it, expect } from 'vitest'
import { createBlankPptx, narrationSpids, openPptx } from '@genoffice/pptx-engine'
import { OP_DOCS, opVocabulary } from '../src/op-docs'
import { runTxn } from '../src/ops/executor'
import '../src/ops/index'

const SILENT_WAV_BASE64 = 'UklGRiwAAABXQVZFZm10IBAAAAABAAEAQB8AAIA+AAACABAAZGF0YQgAAAAAAAAAAAAAAA=='

describe('narration ops', () => {
  it('records and clears a slide narration', async () => {
    const opened = await openPptx(await createBlankPptx())
    const target = { slide: 0 }
    const added = runTxn(opened, {
      ops: [{ op: 'addNarration', target, bytes: SILENT_WAV_BASE64, ext: 'wav', durationMs: 1 }],
    })
    expect(added.failures ?? []).toEqual([])
    expect(narrationSpids(opened.deck.slides[0]!)).toHaveLength(1)
    runTxn(opened, { ops: [{ op: 'clearNarration', target }] })
    expect(narrationSpids(opened.deck.slides[0]!)).toEqual([])
  })

  it('is not offered to the AI', () => {
    for (const name of ['addNarration', 'clearNarration']) {
      expect(OP_DOCS[name]!.aiCallable).toBe(false)
      expect(opVocabulary()).not.toContain(name)
    }
  })
})

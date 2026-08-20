/**
 * Op docs ↔ registry coverage: the doc table is the single source the AI
 * surfaces consume, so it must track the registry exactly — a new op without
 * a doc line (or a doc line for a removed op) fails here, not in production.
 */
import { describe, it, expect } from 'vitest'
import { addElement, createBlankPptx, openPptx } from '@genoffice/pptx-engine'
import { runTxn, opNames } from '../src/main/ops'
import { OP_DOCS, opUsage, opVocabulary } from '../src/shared/op-docs'

// pending: true entries document ops of an in-flight branch ahead of its
// merge so the PRs stay independent; they are hidden from vocabulary and
// usage until registered.
const IN_FLIGHT = new Set(Object.keys(OP_DOCS).filter((n) => OP_DOCS[n]!.pending))

describe('op docs coverage', () => {
  it('every registered op has a doc line', () => {
    const undocumented = opNames().filter((n) => !OP_DOCS[n])
    expect(undocumented).toEqual([])
  })

  it('every doc line matches a registered op', () => {
    const registered = new Set(opNames())
    const stale = Object.keys(OP_DOCS).filter((n) => !registered.has(n) && !IN_FLIGHT.has(n))
    expect(stale).toEqual([])
  })

  it('the vocabulary lists callable ops grouped, and hides byte/clipboard ops', () => {
    const vocab = opVocabulary()
    expect(vocab).toContain('tableMerge')
    expect(vocab).toContain('applyHeaderFooter')
    expect(vocab).toContain('groupElements')
    expect(vocab).not.toContain('addPicture')
    expect(vocab).not.toContain('pasteSlide')
  })

  it('pending ops are hidden from vocabulary and usage until registered', () => {
    for (const name of IN_FLIGHT) {
      expect(opVocabulary()).not.toContain(name)
      expect(opUsage(name)).toBeUndefined()
    }
  })
})

describe('guided errors carry the op signature', () => {
  it('a validation failure appends the one-line usage', async () => {
    const opened = await openPptx(await createBlankPptx())
    const el = addElement(opened.deck.slides[0]!, {
      kind: 'textbox',
      offset: { x: 0, y: 0, cx: 914400, cy: 457200 },
      paragraphs: [{ runs: [{ text: 'T' }] }],
    })
    const r = runTxn(opened, {
      ops: [{ op: 'setFill', target: { slide: 0, el: el.id } }], // fill missing
    })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain(`Usage: setFill ${OP_DOCS.setFill!.sig}`)
  })

  it('an unknown op name still returns the full vocabulary, without a usage line', async () => {
    const opened = await openPptx(await createBlankPptx())
    const r = runTxn(opened, { ops: [{ op: 'sparkle' }] })
    expect(r.applied).toBe(false)
    expect(r.failures![0]!.error).toContain('unknown op "sparkle"')
    expect(r.failures![0]!.error).not.toContain('Usage:')
  })

  it('opUsage returns undefined for unknown names', () => {
    expect(opUsage('sparkle')).toBeUndefined()
  })
})

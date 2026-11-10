/**
 * Op docs ↔ registry coverage: the doc table is the single source the AI
 * surfaces consume, so it must track the registry exactly.
 */
import { describe, expect, it } from 'vitest'
import { opNames, planEditOps } from '../src/renderer/edit-ops'
import { OP_DOCS, opSignatureIndex, opVocabulary } from '../src/shared/op-docs'

const IN_FLIGHT = new Set(Object.keys(OP_DOCS).filter((n) => OP_DOCS[n]!.pending))

describe('op docs coverage', () => {
  it('every registered op has a doc line', () => {
    expect(opNames().filter((n) => !OP_DOCS[n])).toEqual([])
  })

  it('every doc line matches a registered op', () => {
    const registered = new Set(opNames())
    expect(Object.keys(OP_DOCS).filter((n) => !registered.has(n) && !IN_FLIGHT.has(n))).toEqual([])
  })

  it('the vocabulary lists callable ops grouped and hides payload ops', () => {
    const vocab = opVocabulary()
    expect(vocab).toContain('rotatePages')
    expect(vocab).toContain('setFormValue')
    expect(vocab).not.toContain('deleteSavedAnnot')
    expect(vocab).not.toContain('addImageEdit')
    expect(vocab).not.toContain('putTextEdit')
  })

  it('the signature index shows the model-facing signature of every callable op', () => {
    const index = opSignatureIndex()
    for (const n of Object.keys(OP_DOCS)) {
      const d = OP_DOCS[n]!
      if (d.aiCallable === false || d.pending) expect(index).not.toContain(`- ${n} `)
      else expect(index).toContain(`- ${n} ${d.aiSig ?? d.sig}`)
    }
    expect(index).toContain('1-based page number')
    expect(index).not.toContain('pageIndex')
  })

  it('a validation failure appends the one-line usage', () => {
    const plan = planEditOps(
      [{ op: 'rotatePages', pages: [0], dir: 45 }],
      { readOnly: false, pageCount: 2, deleted: new Set(), claimedImages: new Set() },
      () => 'x',
    )
    expect(plan.failures[0]!.error).toContain(`Usage: rotatePages ${OP_DOCS.rotatePages!.sig}`)
  })

  it('an unknown op lists the registry without a usage line', () => {
    const plan = planEditOps(
      [{ op: 'nope' }],
      { readOnly: false, pageCount: 2, deleted: new Set(), claimedImages: new Set() },
      () => 'x',
    )
    expect(plan.failures[0]!.error).toContain('Unknown op "nope"')
    expect(plan.failures[0]!.error).toContain('rotatePages')
    expect(plan.failures[0]!.error).not.toContain('Usage:')
  })
})

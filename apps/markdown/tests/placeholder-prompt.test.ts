import { describe, expect, it } from 'vitest'
import { placeholderInstruction } from '../src/renderer/editor/redact'

/**
 * The prompt this editor sends, pinned.
 *
 * The rules were consolidated into `@genoffice/agent-core` so five editors
 * would stop carrying five near-identical copies. Consolidation is only free if
 * the wording does not move: this editor had its own two consequences — what a
 * marker can stand for here, and what the model must not infer from its
 * absence — and those are the reason a shared assembler takes a spec rather
 * than a fixed prompt.
 *
 * So the whole string is a golden here, not a summary of it. A change to the
 * shared rules, to this spec, or to either editor's consequences shows up here
 * as a diff, which is the only place a change of this kind should ever be
 * visible.
 */

const LABELS = ['client phone', 'API key', 'client phone', '  ']

const EXPECTED =
  '## Private placeholders\nThis document contains {{...}} placeholders. Each one stands in for text the reader has deliberately withheld from you; you cannot see what is inside and that is the point.\n\nTreat every placeholder as one indivisible object:\n- Copy it character for character — same letters, same order, same spacing.\n- Never split it across a line break or put a space inside it.\n- Never merge two into one, never split one into several, never reorder them.\n- Never rename, translate, re-case, expand or shorten it.\n- Never drop one, and never add a placeholder that was not already there.\n\nWrite the prose around them as if each stood for the words it replaces, so "call {{client phone}}" reads as a natural instruction to phone someone.\nIf a request needs what a placeholder hides, write around it rather than guessing.\n\nThe placeholders in this document:\n- {{client phone}}\n- {{API key}}\n- {{private}}'

describe('the placeholder prompt', () => {
  it('is exactly what this editor has always sent', () => {
    expect(placeholderInstruction(LABELS)).toBe(EXPECTED)
  })

  it('and for an empty label set', () => {
    expect(placeholderInstruction([])).toBe(placeholderInstruction([]).replace(EXPECTED, EXPECTED))
    expect(placeholderInstruction([])).toContain('## Private placeholders')
  })

  it('keeps its own wording, worked example included, in a Latin script', () => {
    // The reader's labels may be in any script, but this instruction ships to
    // every model in every language, and a worked example written in one
    // script teaches that script's conventions rather than the rule. With no
    // labels passed, every character in here is ours to keep Latin.
    const text = placeholderInstruction([])
    expect(text).not.toMatch(/\p{Script=Han}/u)
    expect(text).toContain('call {{client phone}}')
  })
})

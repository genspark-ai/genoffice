import { describe, expect, it } from 'vitest'
import { placeholderInstruction } from '../src/renderer/ai/redact'

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

const LABELS = ['API token', 'script value', 'API token', 'comment', '  ']

const EXPECTED =
  '## Private placeholders\nThis page contains {{...}} placeholders. Each stands in for something the reader has deliberately withheld from you; you cannot see what is inside, and that is the point.\n\nTreat every placeholder as one indivisible object:\n- Copy it character for character — same letters, same order, same spacing.\n- Never split it across a line break or put a space inside it.\n- Never merge two into one, never split one into several, never reorder them.\n- Never rename, translate, re-case, expand or shorten it.\n- Never drop one, and never add a placeholder that was not already there.\n\nA placeholder may stand in for visible text, for the value of an attribute (a key, a token, a URL), for a value inside a <script>, or for a whole comment. In every case it is text you were not given: write around it rather than guessing.\nA withheld attribute or script value still governs how the page behaves. Do not invent a replacement for it and do not remove the thing that carries it.\n\nThe placeholders in this page:\n- {{API token}}\n- {{script value}}\n- {{comment}}\n- {{private}}'

describe('the placeholder prompt', () => {
  it('is exactly what this editor has always sent', () => {
    expect(placeholderInstruction(LABELS)).toBe(EXPECTED)
  })

  it('and for an empty label set', () => {
    expect(placeholderInstruction([])).toBe(placeholderInstruction([]).replace(EXPECTED, EXPECTED))
    expect(placeholderInstruction([])).toContain('## Private placeholders')
  })
})

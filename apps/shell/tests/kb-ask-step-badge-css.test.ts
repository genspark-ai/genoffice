import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

// The "Ask your knowledge base" onboarding step badges (issue #2385): the
// number sits in a blue circle. A generic `.kb-ask-step span` rule used to
// override the badge's `display: flex` (higher specificity), so the digit
// rendered left-aligned outside the circle instead of centered. Keep the
// description rule scoped to the text wrapper, and keep the badge centered.
const css = readFileSync(resolve(__dirname, '../src/renderer/src/home.css'), 'utf8')

describe('KbAsk step badge CSS', () => {
  it('keeps the badge flex-centered and the description rule off the badge', () => {
    const badge = /\.kb-ask-step-num\s*\{([^}]*)\}/.exec(css)
    expect(badge, 'missing .kb-ask-step-num rule').toBeTruthy()
    expect(badge![1]).toContain('display: flex')
    expect(badge![1]).toContain('align-items: center')
    expect(badge![1]).toContain('justify-content: center')
    // a bare descendant selector would match the badge span again
    expect(css).not.toMatch(/\.kb-ask-step\s+span\s*\{/)
    expect(css).toContain('.kb-ask-step > div > span')
  })
})

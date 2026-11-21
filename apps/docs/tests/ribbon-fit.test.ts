import { describe, expect, it } from 'vitest'
import { fitSteps } from '../src/renderer/components/ribbon-fit'

describe('fitSteps', () => {
  const stack = (name: string, data: { wide?: string; collapse?: string } = {}) => {
    const el = document.createElement('div')
    el.id = name
    Object.assign(el.dataset, data)
    return el
  }
  const names = (stacks: HTMLElement[]) => fitSteps(stacks).map((s) => `${s.stack.id}>${s.to}`)

  it('takes the ranked stacks first, then the rest from the right', () => {
    const stacks = [
      stack('pages'),
      stack('links'),
      stack('text', { collapse: '1' }),
      stack('symbols'),
    ]
    expect(names(stacks)).toEqual(['text>icons', 'symbols>icons', 'links>icons', 'pages>icons'])
  })

  it('folds buttons into a stack before any stack of that rank loses its captions', () => {
    const stacks = [
      stack('pages', { wide: '2' }),
      stack('links', { wide: '1' }),
      stack('text', { collapse: '2' }),
    ]
    expect(names(stacks)).toEqual([
      'links>stack',
      'pages>stack',
      'text>icons',
      'links>icons',
      'pages>icons',
    ])
  })

  it('a stack is a stack before it is icons, whatever its ranks say', () => {
    expect(names([stack('proofing', { wide: '3', collapse: '1' })])).toEqual([
      'proofing>stack',
      'proofing>icons',
    ])
  })
})

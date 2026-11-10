/**
 * PDF toast live-region + outline empty-state/tree semantics.
 *
 * OutlinePanel mounts directly with the dialog-a11y.test.ts createRoot/act
 * harness. The App toasts live inside the full App component (too heavy to
 * mount here), so they are verified with a static source assertion that both
 * `.pdf-toast` containers carry role="status" (docs toast.tsx pattern).
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { afterEach, beforeAll, describe, expect, it } from 'vitest'
import { OutlinePanel, type OutlineNode } from '../src/renderer/OutlinePanel'

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
})

async function renderPanel(props: {
  outline: OutlineNode[]
  note?: string
  label?: string
  emptyLabel?: string
  currentDest?: unknown
}) {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(createElement(OutlinePanel, { ...props, onGoToDest: () => {} }))
    await Promise.resolve()
  })
  return container
}

const sample: OutlineNode[] = [
  {
    title: 'Chapter 1',
    dest: 'd1',
    items: [{ title: 'Section 1.1', dest: 'd1.1' }],
  },
  { title: 'Chapter 2', dest: 'd2' },
]

describe('OutlinePanel a11y', () => {
  it('exposes a labelled tree', async () => {
    const el = await renderPanel({
      outline: sample,
      label: 'Outline',
      emptyLabel: 'No results',
    })
    const tree = el.querySelector('[role="tree"]')
    expect(tree).not.toBeNull()
    expect(tree!.getAttribute('aria-label')).toBe('Outline')
  })

  it('marks items as treeitems with level and expanded state', async () => {
    const el = await renderPanel({
      outline: sample,
      label: 'Outline',
      emptyLabel: 'No results',
    })
    const items = [...el.querySelectorAll('[role="treeitem"]')]
    expect(items).toHaveLength(3)
    expect(items[0].getAttribute('aria-level')).toBe('1')
    expect(items[0].getAttribute('aria-expanded')).toBe('true')
    expect(items[1].getAttribute('aria-level')).toBe('2')
    expect(items[1].hasAttribute('aria-expanded')).toBe(false)
  })

  it('marks the current destination with aria-current', async () => {
    const el = await renderPanel({
      outline: sample,
      label: 'Outline',
      emptyLabel: 'No results',
      currentDest: 'd2',
    })
    const items = [...el.querySelectorAll('[role="treeitem"]')]
    const current = items.filter((n) => n.getAttribute('aria-current') === 'true')
    expect(current).toHaveLength(1)
    expect(current[0].textContent).toBe('Chapter 2')
  })

  it('caps nesting depth instead of recursing into hostile outlines', async () => {
    let deep: OutlineNode = { title: 'leaf', dest: 'leaf' }
    for (let i = 0; i < 200; i++) deep = { title: `level ${i}`, dest: `d${i}`, items: [deep] }
    const el = await renderPanel({ outline: [deep], label: 'Outline', emptyLabel: 'No results' })
    const items = [...el.querySelectorAll('[role="treeitem"]')]
    expect(items.length).toBeGreaterThan(0)
    expect(items.length).toBeLessThanOrEqual(33)
  })

  it('renders the empty state as a live region reusing the no-results copy', async () => {
    const el = await renderPanel({
      outline: [],
      label: 'Outline',
      emptyLabel: 'No results',
    })
    expect(el.querySelector('[role="treeitem"]')).toBeNull()
    const empty = el.querySelector('.pdf-outline-empty[role="status"]')
    expect(empty).not.toBeNull()
    expect(empty!.textContent).toBe('No results')
  })
})

describe('pdf toast live region', () => {
  it('gives both toast containers role=status', () => {
    const src = readFileSync(join(__dirname, '..', 'src', 'renderer', 'App.tsx'), 'utf8')
    expect(src).toContain('<div className="pdf-toast" role="status">')
    expect(src).toContain('<div className="pdf-toast pdf-toast-notice" role="status">')
  })
})

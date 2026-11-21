// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import type { KbFileInfo } from '@genoffice/agent-core'
import { KbCorpusList } from '../src/renderer/src/KbAskPanel'
import { LocaleProvider } from '../src/renderer/src/locale'

/**
 * The files behind the Ask panel's knowledge-base count.
 *
 * The panel already loaded the starred corpus into a ref for the knowledge-base
 * skill and kept only the number in state, so the reader could see how many
 * files the AI could answer from but never which ones. These tests pin the
 * list that now opens from the header.
 */

const corpus: KbFileInfo[] = [
  {
    path: '/Users/me/contracts/lease.docx',
    folder: 'contracts',
    name: 'lease.docx',
    ext: 'docx',
    sizeBytes: 2048,
  },
  {
    path: '/Users/me/budgets/q3.xlsx',
    folder: 'budgets',
    name: 'q3.xlsx',
    ext: 'xlsx',
    sizeBytes: 512000,
  },
  {
    path: '/Users/me/notes/readme.txt',
    folder: '',
    name: 'readme.txt',
    ext: 'txt',
    sizeBytes: 120,
  },
]

let root: Root
let container: HTMLDivElement
const opened: string[] = []

beforeEach(() => {
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  opened.length = 0
  container = document.createElement('div')
  document.body.append(container)
  root = createRoot(container)
})

afterEach(async () => {
  await act(async () => root.unmount())
  container.remove()
  vi.unstubAllGlobals()
})

async function mount(files: readonly KbFileInfo[]): Promise<void> {
  await act(async () => {
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(KbCorpusList, {
          files,
          onOpenPath: (path: string) => {
            opened.push(path)
          },
          label: 'Starred files in the knowledge base',
        }),
      }),
    )
    await new Promise((resolve) => setTimeout(resolve, 0))
  })
}

function rows(): HTMLButtonElement[] {
  return [...container.querySelectorAll<HTMLButtonElement>('.kb-ask-corpus-row')]
}
function cell(row: HTMLButtonElement, cls: string): string | undefined {
  return row.querySelector('.' + cls)?.textContent ?? undefined
}

describe('the knowledge-base corpus list', () => {
  it('lists every starred file with its folder and size', async () => {
    await mount(corpus)
    expect(rows()).toHaveLength(3)
    expect(rows().map((row) => cell(row, 'kb-ask-corpus-name'))).toEqual([
      'lease.docx',
      'q3.xlsx',
      'readme.txt',
    ])
    expect(rows().map((row) => cell(row, 'kb-ask-corpus-folder'))).toEqual([
      'contracts',
      'budgets',
      undefined,
    ])
    expect(rows().map((row) => cell(row, 'kb-ask-corpus-size'))).toEqual(['2KB', '500KB', '120B'])
  })

  it('gives every row the full path as its tooltip', async () => {
    await mount(corpus)
    expect(rows().map((row) => row.getAttribute('title'))).toEqual(corpus.map((file) => file.path))
  })

  it('opens the file a row names, the way a citation does', async () => {
    await mount(corpus)
    await act(async () => rows()[1]!.click())
    expect(opened).toEqual(['/Users/me/budgets/q3.xlsx'])
  })

  it('keeps the rows keyboard-reachable native buttons', async () => {
    await mount(corpus)
    expect(rows().every((row) => row.tagName === 'BUTTON')).toBe(true)
  })

  it('labels the list for assistive technology', async () => {
    await mount(corpus)
    expect(container.querySelector('.kb-ask-corpus')?.getAttribute('role')).toBe('list')
    expect(container.querySelector('.kb-ask-corpus')?.getAttribute('aria-label')).toBe(
      'Starred files in the knowledge base',
    )
    expect(container.querySelectorAll('[role="listitem"]')).toHaveLength(3)
  })
})

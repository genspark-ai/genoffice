/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import type {
  FileSearchHit,
  HomeApi,
  InstalledPlugin,
  RecentEntry,
  UserSkillEntry,
} from '../src/shared/home-api'
import type { PluginComponents } from '../src/shared/found-skill'
import { CommandPalette, folderOf, scoreMatch } from '../src/renderer/src/CommandPalette'
import {
  createAskAiProvider,
  createFilesProvider,
  defaultPaletteProviders,
  mergePaletteRows,
  type PaletteContext,
  type PaletteProvider,
  type PaletteRow,
} from '../src/renderer/src/palette-providers'
import { COMMAND_PALETTE_STRINGS } from '../src/renderer/src/command-palette-strings'

/**
 * The Ctrl/Cmd+K launcher. Everything it does goes through window.aiOffice:
 * the lists it reads on open, and the bridge action Enter fires. Esc and
 * backdrop clicks only ever call onClose.
 */

const actEnvironment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
actEnvironment.IS_REACT_ACT_ENVIRONMENT = true

let host: HTMLDivElement
let root: Root
let closed: number
let openedPaths: string[]
let created: string[]
let skillsSettingsOpened: number
let modelSettingsOpened: number
let pluginToggles: Array<[string, boolean]>
let askedQueries: string[]

const recent = (over: Partial<RecentEntry>): RecentEntry => ({
  path: 'C:\\Docs\\report.docx',
  name: 'report.docx',
  ext: 'docx',
  mtimeMs: 1,
  sizeBytes: 1,
  starred: false,
  ...over,
})

const hit = (over: Partial<FileSearchHit>): FileSearchHit => ({
  path: 'C:\\Docs\\hit.docx',
  name: 'hit.docx',
  ext: 'docx',
  mtimeMs: 1,
  sizeBytes: 1,
  starred: false,
  snippet: null,
  needles: [],
  ...over,
})

/** a provider context for driving the registry without mounting the palette */
const ctx = (over: Partial<PaletteContext> = {}): PaletteContext => ({
  strings: COMMAND_PALETTE_STRINGS.en,
  recents: [],
  skills: [],
  plugins: [],
  onAskAi: () => {},
  ...over,
})

const skill = (over: Partial<UserSkillEntry>): UserSkillEntry => ({
  name: 'pptx-builder',
  description: 'Build a deck',
  relevant: true,
  ...over,
})

const plugin = (
  over: Partial<InstalledPlugin> & { components?: Partial<PluginComponents> },
): InstalledPlugin => ({
  name: 'deck-tools',
  version: '1.0.0',
  description: 'Helpers for decks',
  author: 'someone',
  enabled: true,
  components: { skills: [], commands: [], mcpServers: [], ...over.components },
  format: 'claude',
  ...over,
})

beforeEach(() => {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
  closed = 0
  openedPaths = []
  created = []
  skillsSettingsOpened = 0
  modelSettingsOpened = 0
  pluginToggles = []
  askedQueries = []
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
})

function mountApi(
  recentsList: RecentEntry[],
  skills: UserSkillEntry[] = [],
  plugins: InstalledPlugin[] = [],
): void {
  window.aiOffice = {
    recents: async () => ({
      entries: recentsList,
      total: recentsList.length,
      totalAll: recentsList.length,
    }),
    skillsList: async () => skills,
    pluginsList: async () => plugins,
    openPath: async (path: string) => {
      openedPaths.push(path)
    },
    browse: async () => {
      created.push('browse')
    },
    newDoc: async () => {
      created.push('doc')
    },
    newSheet: async () => {
      created.push('sheet')
    },
    newSlide: async () => {
      created.push('slide')
    },
    newMarkdown: async () => {
      created.push('markdown')
    },
    newHtml: async () => {
      created.push('html')
    },
    newPdf: async () => {
      created.push('pdf')
    },
    openSkillsSettings: async () => {
      skillsSettingsOpened += 1
    },
    openAiModelSettings: async () => {
      modelSettingsOpened += 1
    },
    pluginsSetEnabled: async (name: string, enabled: boolean) => {
      pluginToggles.push([name, enabled])
    },
  } as unknown as HomeApi
}

async function render(providers?: PaletteProvider[]): Promise<void> {
  await act(async () => {
    root.render(
      createElement(CommandPalette, {
        lang: 'en',
        onClose: () => (closed += 1),
        onAskAi: (q) => askedQueries.push(q),
        ...(providers ? { providers } : {}),
      }),
    )
  })
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
  })
}

/** the palette's search input */
const input = (): HTMLInputElement => host.querySelector<HTMLInputElement>('.cmd-input')

/** the flat result rows, in visual order */
const rows = (): HTMLButtonElement[] => [...host.querySelectorAll<HTMLButtonElement>('.cmd-row')]

const activeRow = (): HTMLButtonElement | undefined =>
  host.querySelector<HTMLButtonElement>('.cmd-row.active')

const rowTitles = (): string[] =>
  rows().map((row) => row.querySelector('.cmd-row-title')?.textContent)

/** type into the search input through the native setter, so React's own value
 *  tracker registers the change and onChange fires */
function typeInto(text: string): void {
  const el = input()
  expect(el).toBeTruthy()
  const proto = window.HTMLInputElement.prototype
  Object.getOwnPropertyDescriptor(proto, 'value')?.set?.call(el, text)
  el.dispatchEvent(new Event('input', { bubbles: true }))
}

async function actType(text: string): Promise<void> {
  await act(async () => {
    typeInto(text)
    await Promise.resolve()
    await Promise.resolve()
  })
}

/** one keydown on the search input, as the keyboard driver sends it */
async function press(key: string): Promise<void> {
  await act(async () => {
    input().dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true }))
    await Promise.resolve()
    await Promise.resolve()
  })
}

async function clickRow(title: string): Promise<void> {
  const row = rows().find((r) => r.textContent?.includes(title))
  expect(row).toBeTruthy()
  await act(async () => {
    row!.click()
    await Promise.resolve()
    await Promise.resolve()
  })
}

describe('scoreMatch', () => {
  it('ranks a title prefix over a title substring, then a detail hit, then a subsequence', () => {
    const prefix = scoreMatch('sheet', 'sheet-maker', 'a skill')
    const inside = scoreMatch('sheet', 'New spreadsheet', '')
    const detail = scoreMatch('sheet', 'New file', 'AI Sheets · xlsx')
    const fuzzy = scoreMatch('nsp', 'New spreadsheet', '')
    expect(prefix).not.toBeNull()
    expect(inside).not.toBeNull()
    expect(detail).not.toBeNull()
    expect(fuzzy).not.toBeNull()
    expect(prefix!).toBeGreaterThan(inside!)
    expect(inside!).toBeGreaterThan(detail!)
    expect(detail!).toBeGreaterThan(fuzzy!)
  })

  it('matches case-insensitively and as a subsequence over the title', () => {
    expect(scoreMatch('SHEET', 'New spreadsheet', 'AI Sheets · xlsx')).not.toBeNull()
    // letters in order, not contiguous
    expect(scoreMatch('nssh', 'New spreadsheet', '')).not.toBeNull()
    // a letter missing breaks it
    expect(scoreMatch('nxw', 'New spreadsheet', '')).toBeNull()
  })

  it('scores every row as a match for an empty query', () => {
    expect(scoreMatch('', 'anything')).toBe(0)
    expect(scoreMatch('   ', 'anything')).toBe(0)
  })

  it('never matches text that lacks a character of the query', () => {
    expect(scoreMatch('表格', '新建文档', 'Word · docx')).toBeNull()
    expect(scoreMatch('表格', '新建表格', 'Sheet · xlsx')).not.toBeNull()
  })

  it('keeps CJK and mixed-script queries working through the detail line', () => {
    // typing "sheet" while the UI is Chinese still finds the spreadsheet action
    expect(scoreMatch('sheet', '新建表格', 'AI Sheets · xlsx')).not.toBeNull()
    expect(scoreMatch('表格', '新建表格', 'AI Sheets · xlsx')).not.toBeNull()
  })
})

describe('CommandPalette', () => {
  it('focuses the input on open', async () => {
    mountApi([])
    await render()
    expect(document.activeElement).toBe(input())
  })

  it('shows the create actions and the top 5 recents before anything is typed', async () => {
    mountApi(
      Array.from({ length: 8 }, (_, i) =>
        recent({ path: `C:\\Docs\\file-${i}.docx`, name: `file-${i}.docx` }),
      ),
    )
    await render()
    // the fixed create actions, in order
    expect(rowTitles().slice(0, 7)).toEqual([
      'New document',
      'New spreadsheet',
      'New presentation',
      'New Markdown',
      'New web page',
      'New PDF',
      'Open a local file…',
    ])
    // only the five newest recents ride along
    expect(rows().slice(7)).toHaveLength(5)
    expect(rowTitles()[7]).toBe('file-0.docx')
    expect(rowTitles()[11]).toBe('file-4.docx')
    // no skills, settings or plugins groups on an empty query
    expect(host.querySelector('.cmd-group-title')?.textContent).toBe('New')
    expect(host.querySelectorAll('.cmd-group-title')).toHaveLength(2)
  })

  it('typing filters the recents by file name', async () => {
    mountApi([
      recent({ path: 'C:\\Docs\\budget-2026.xlsx', name: 'budget-2026.xlsx', ext: 'xlsx' }),
      recent({ path: 'C:\\Docs\\report.docx', name: 'report.docx' }),
    ])
    await render()
    await actType('budget')
    // the recent match, then the natural-language fallback that never leaves a
    // non-empty query unanswered
    expect(rowTitles()).toEqual(['budget-2026.xlsx', 'Ask AI: budget'])
    // the folder rides on the row as its detail line
    expect(rows()[0]!.querySelector('.cmd-row-detail')?.textContent).toBe('C:\\Docs')
  })

  it('Enter opens the selected recent through openPath and closes', async () => {
    mountApi([
      recent({ path: 'C:\\Docs\\budget-2026.xlsx', name: 'budget-2026.xlsx', ext: 'xlsx' }),
    ])
    await render()
    await actType('budget')
    await press('Enter')
    expect(openedPaths).toEqual(['C:\\Docs\\budget-2026.xlsx'])
    expect(closed).toBe(1)
    expect(created).toEqual([])
  })

  it('Esc closes without running anything', async () => {
    mountApi([recent({})])
    await render()
    await actType('report')
    await press('Escape')
    expect(closed).toBe(1)
    expect(openedPaths).toEqual([])
    expect(created).toEqual([])
  })

  it('a backdrop click closes too', async () => {
    mountApi([recent({})])
    await render()
    const overlay = host.querySelector<HTMLElement>('.cmd-overlay')
    await act(async () => {
      overlay!.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }))
      await Promise.resolve()
    })
    expect(closed).toBe(1)
  })

  it('keyboard navigation moves and wraps over the whole flat list', async () => {
    mountApi(
      Array.from({ length: 3 }, (_, i) =>
        recent({ name: `file-${i}.docx`, path: `C:\\D\\file-${i}.docx` }),
      ),
    )
    await render()
    // 7 create actions + 3 recents = 10 rows, the first is active
    expect(rows()).toHaveLength(10)
    expect(activeRow()?.textContent).toContain('New document')
    // up from the first wraps to the last row (the third recent)
    await press('ArrowUp')
    expect(activeRow()?.textContent).toContain('file-2.docx')
    // up again steps within the recents
    await press('ArrowUp')
    expect(activeRow()?.textContent).toContain('file-1.docx')
    // down twice returns to the first create action, wrapping through the end
    await press('ArrowDown')
    await press('ArrowDown')
    expect(activeRow()?.textContent).toContain('New document')
    // the highlight follows the combobox's aria-activedescendant
    const activeIndex = rows().findIndex((row) => row === activeRow())
    expect(input().getAttribute('aria-activedescendant')).toBe(`cmd-opt-${activeIndex}`)
  })

  it('a query with no matches falls through to the Ask AI row, and Enter submits it', async () => {
    mountApi([recent({})])
    await render()
    await actType('zzz-no-match')
    // no empty state anymore: the fallback row is the one row that always answers
    expect(rows()).toHaveLength(1)
    expect(rowTitles()).toEqual(['Ask AI: zzz-no-match'])
    expect(host.querySelector('.cmd-empty')).toBeFalsy()
    // arrow keys cycle the single row and Enter hands the query to the host
    await press('ArrowDown')
    expect(activeRow()?.textContent).toContain('Ask AI')
    await press('Enter')
    expect(askedQueries).toEqual(['zzz-no-match'])
    expect(closed).toBe(1)
    expect(openedPaths).toEqual([])
  })

  it('lists installed skills and plugin components under Skills & commands', async () => {
    mountApi(
      [],
      [skill({ name: 'notes-report', description: 'Turn notes into a report' })],
      [plugin({ name: 'deck-tools', components: { skills: ['builder'], commands: ['deploy'] } })],
    )
    await render()
    await actType('notes')
    expect(host.textContent).toContain('Skills & commands')
    expect(rowTitles()).toEqual(['notes-report', 'Ask AI: notes'])
    // a plugin's components surface as plugin:name entries, searched by name
    await actType('deck-tools:deploy')
    expect(rowTitles()).toEqual(['deck-tools:deploy', 'Ask AI: deck-tools:deploy'])
    expect(rows()[0]!.querySelector('.cmd-row-detail')?.textContent).toBe('Helpers for decks')
  })

  it('picking a skill opens the skills settings page (the ask view is not bridge-reachable)', async () => {
    mountApi([], [skill({ name: 'notes-report' })])
    await render()
    await actType('notes')
    await press('Enter')
    expect(skillsSettingsOpened).toBe(1)
    expect(closed).toBe(1)
  })

  it('typing sheet offers the spreadsheet action, and Enter runs newSheet', async () => {
    mountApi([])
    await render()
    await actType('sheet')
    expect(rowTitles()).toContain('New spreadsheet')
    // navigate onto it: it is the top match for "sheet"
    expect(activeRow()?.textContent).toContain('New spreadsheet')
    await press('Enter')
    expect(created).toEqual(['sheet'])
    expect(closed).toBe(1)
  })

  it('jumping to the AI model settings goes through the bridge', async () => {
    mountApi([])
    await render()
    await actType('model')
    expect(rowTitles()).toContain('AI model settings')
    await press('Enter')
    expect(modelSettingsOpened).toBe(1)
    expect(closed).toBe(1)
    expect(skillsSettingsOpened).toBe(0)
  })

  it('plugins list with a Disable chip, and Enter toggles them through the bridge', async () => {
    mountApi([], [], [plugin({ enabled: true }), plugin({ name: 'off-one', enabled: false })])
    await render()
    await actType('deck-tools')
    expect(rowTitles()).toEqual(['deck-tools', 'Ask AI: deck-tools'])
    expect(rows()[0]!.querySelector('.cmd-chip')?.textContent).toBe('Disable')
    await press('Enter')
    expect(pluginToggles).toEqual([['deck-tools', false]])
    expect(closed).toBe(1)
    // a disabled plugin offers the inverse action
    await actType('off-one')
    expect(rows()[0]!.querySelector('.cmd-chip')?.textContent).toBe('Enable')
  })

  it('clicking a row runs it, and hovering moves the highlight', async () => {
    mountApi([
      recent({ path: 'C:\\Docs\\a.docx', name: 'a.docx' }),
      recent({ path: 'C:\\Docs\\b.docx', name: 'b.docx' }),
    ])
    await render()
    expect(activeRow()?.textContent).toContain('New document')
    await act(async () => {
      rows()[8]!.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }))
      await Promise.resolve()
    })
    expect(activeRow()?.textContent).toContain('b.docx')
    await clickRow('b.docx')
    expect(openedPaths).toEqual(['C:\\Docs\\b.docx'])
    expect(closed).toBe(1)
  })

  it('the Ask AI fallback never shows on an empty query', async () => {
    mountApi([recent({})])
    await render()
    expect(rowTitles().some((title) => title.startsWith('Ask AI'))).toBe(false)
    expect(host.querySelector('.cmd-empty')).toBeFalsy()
  })

  it('the fallback row carries its hint as the detail line', async () => {
    mountApi([recent({})])
    await render()
    await actType('zzz-no-match')
    expect(rowTitles()).toEqual(['Ask AI: zzz-no-match'])
    expect(rows()[0]!.querySelector('.cmd-row-detail')?.textContent).toBe(
      'Send this question to the Ask AI panel',
    )
  })

  it('dims a missing recent but Enter still tries to open it', async () => {
    mountApi([recent({ name: 'gone.docx', path: 'C:\\Docs\\gone.docx', missing: true })])
    await render()
    await actType('gone')
    expect(rows()[0]!.className).toContain('missing')
    expect(rows()[0]!.querySelector('.cmd-row-detail')?.textContent).toBe('File unavailable')
    await press('Enter')
    expect(openedPaths).toEqual(['C:\\Docs\\gone.docx'])
  })

  it('caps the result list at 50 rows in the fixed group order', async () => {
    mountApi(
      Array.from({ length: 80 }, (_, i) =>
        recent({ path: `C:\\Docs\\num-${i}.docx`, name: `num-${i}.docx` }),
      ),
      [skill({ name: 'num-skill', description: '' })],
    )
    await render()
    await actType('num')
    expect(rows()).toHaveLength(50)
    // the create group leads, then recents; the skill lands after 7 + 43 rows
    expect(rowTitles()[0]).toContain('New')
    expect(host.querySelectorAll('.cmd-group-title')[0]!.textContent).toBe('New')
    expect(host.querySelectorAll('.cmd-group-title')[1]!.textContent).toBe('Recent files')
    expect(rowTitles().includes('num-skill')).toBe(false)
  })

  it('survives a preload without the palette read methods', async () => {
    window.aiOffice = {} as unknown as HomeApi
    await render()
    await actType('doc')
    // the create actions still work; the sources simply read as empty
    expect(rowTitles()).toContain('New document')
    expect(host.querySelector('.cmd-empty')).toBeFalsy()
  })
})

describe('folderOf', () => {
  it('takes the directory part, both separators honored', () => {
    expect(folderOf('C:\\Docs\\report.docx')).toBe('C:\\Docs')
    expect(folderOf('/home/notes.md')).toBe('/home')
    expect(folderOf('report.docx')).toBe('')
    expect(folderOf('\\report.docx')).toBe('')
  })
})

// ── the provider registry ──

const stubRow = (group: PaletteRow['group'], id: string, path?: string): PaletteRow => ({
  id,
  group,
  title: id,
  detail: '',
  icon: { kind: 'pen' },
  missing: false,
  chip: '',
  path,
  run: () => {},
})

describe('defaultPaletteProviders', () => {
  it('registers every source in a stable priority order', () => {
    const providers = defaultPaletteProviders()
    expect(providers.map((p) => p.id)).toEqual([
      'create',
      'recent',
      'files',
      'skills',
      'settings',
      'plugins',
      'ask',
    ])
    const priorities = providers.map((p) => p.priority)
    expect([...priorities].sort((a, b) => a - b)).toEqual(priorities)
  })

  it('reproduces the legacy empty-query list: the create actions plus the five newest recents', () => {
    const recents = Array.from({ length: 8 }, (_, i) =>
      recent({ path: `C:\\Docs\\file-${i}.docx`, name: `file-${i}.docx` }),
    )
    const context = ctx({ recents })
    const parts = defaultPaletteProviders().map((p) => ({
      priority: p.priority,
      rows: p.getRows('', context) as PaletteRow[],
    }))
    const rows = mergePaletteRows(parts)
    expect(rows.map((r) => r.group)).toEqual([
      ...Array(7).fill('create'),
      ...Array.from({ length: 5 }, () => 'recent'),
    ])
    expect(rows[7]!.title).toBe('file-0.docx')
  })
})

describe('createAskAiProvider', () => {
  it('is empty for a blank query and a single trailing row otherwise', () => {
    const provider = createAskAiProvider()
    expect(provider.getRows('', ctx())).toEqual([])
    expect(provider.getRows('   ', ctx())).toEqual([])
    const rows = provider.getRows('  summarize Q3  ', ctx()) as PaletteRow[]
    expect(rows).toHaveLength(1)
    expect(rows[0]!.group).toBe('ask')
    expect(rows[0]!.title).toBe('Ask AI: summarize Q3')
  })

  it('runs by handing the trimmed query to the host', () => {
    const asked: string[] = []
    const provider = createAskAiProvider()
    const rows = provider.getRows('hello', ctx({ onAskAi: (q) => asked.push(q) })) as PaletteRow[]
    rows[0]!.run()
    expect(asked).toEqual(['hello'])
  })
})

describe('createFilesProvider', () => {
  const page = (hits: FileSearchHit[]) => ({
    hits,
    total: hits.length,
    index: { indexed: 1, pending: 0, scanning: false },
  })

  it('answers [] synchronously for a blank query and never searches', () => {
    let calls = 0
    const provider = createFilesProvider({
      search: async () => {
        calls += 1
        return page([])
      },
    })
    expect(provider.getRows('', ctx())).toEqual([])
    expect(provider.getRows('   ', ctx())).toEqual([])
    expect(calls).toBe(0)
  })

  it('debounces the search, caps the rows and maps hits to openable file rows', async () => {
    const searched: string[] = []
    const opened: string[] = []
    const provider = createFilesProvider({
      debounceMs: 0,
      search: async (q) => {
        searched.push(q)
        return page(
          Array.from({ length: 8 }, (_, i) =>
            hit({ path: `C:\\Docs\\hit-${i}.docx`, name: `hit-${i}.docx` }),
          ),
        )
      },
      open: (path) => opened.push(path),
    })
    const rows = await provider.getRows('hit', ctx())
    expect(searched).toEqual(['hit'])
    // at most five file rows ride along
    expect(rows).toHaveLength(5)
    expect(rows[0]!.group).toBe('files')
    expect(rows[0]!.detail).toBe('C:\\Docs')
    rows[0]!.run()
    expect(opened).toEqual(['C:\\Docs\\hit-0.docx'])
    // the same query again reuses the settled answer instead of re-searching
    const again = await provider.getRows('hit', ctx())
    expect(searched).toEqual(['hit'])
    expect(again.map((r) => r.id)).toEqual(rows.map((r) => r.id))
  })

  it('settles to no rows when the search fails, so other sources still answer', async () => {
    const provider = createFilesProvider({
      debounceMs: 0,
      search: () => Promise.reject(new Error('index busy')),
    })
    await expect(provider.getRows('hit', ctx())).resolves.toEqual([])
  })

  it('a superseded query resolves to no rows and its search never fires', async () => {
    const searched: string[] = []
    const provider = createFilesProvider({
      debounceMs: 5,
      search: async (q) => {
        searched.push(q)
        return page([hit({})])
      },
    })
    const first = provider.getRows('old', ctx()) as Promise<PaletteRow[]>
    const second = provider.getRows('new', ctx()) as Promise<PaletteRow[]>
    await expect(first).resolves.toEqual([])
    await expect(second).resolves.toHaveLength(1)
    expect(searched).toEqual(['new'])
  })
})

describe('mergePaletteRows', () => {
  it('orders groups by priority regardless of contribution order', () => {
    const rows = mergePaletteRows([
      { priority: 100, rows: [stubRow('ask', 'ask:ai')] },
      { priority: 10, rows: [stubRow('create', 'create:doc')] },
      { priority: 20, rows: [stubRow('recent', 'recent:x', '/x')] },
    ])
    expect(rows.map((r) => r.id)).toEqual(['create:doc', 'recent:x', 'ask:ai'])
  })

  it('a file-index hit shadows the recent row of the same path', () => {
    // both rows open the same file; the hit matched the live query, the recent
    // only repeats an old open, so the recent yields its slot
    const rows = mergePaletteRows([
      {
        priority: 20,
        rows: [
          stubRow('recent', 'recent:/a.docx', '/a.docx'),
          stubRow('recent', 'recent:/b.docx', '/b.docx'),
        ],
      },
      { priority: 30, rows: [stubRow('files', 'file:/a.docx', '/a.docx')] },
    ])
    expect(rows.map((r) => r.id)).toEqual(['recent:/b.docx', 'file:/a.docx'])
  })

  it('caps the list at 50 rows and always keeps the Ask AI fallback last', () => {
    const many = Array.from({ length: 80 }, (_, i) => stubRow('recent', `recent:${i}`))
    const rows = mergePaletteRows([
      { priority: 20, rows: many },
      { priority: 100, rows: [stubRow('ask', 'ask:ai')] },
    ])
    expect(rows).toHaveLength(50)
    expect(rows[49]!.id).toBe('ask:ai')
  })
})

describe('CommandPalette + FilesProvider (mounted)', () => {
  /** the default registry with the files source swapped for a test double */
  const providersWithFiles = (
    search: (q: string) => Promise<{
      hits: FileSearchHit[]
      total: number
      index: { indexed: number; pending: number; scanning: boolean }
    }>,
  ): PaletteProvider[] =>
    defaultPaletteProviders().map((p) =>
      p.id === 'files' ? createFilesProvider({ debounceMs: 0, search }) : p,
    )

  /** the debounced search lands a macrotask later even with debounceMs 0 */
  const flushSearch = () =>
    act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10))
      await Promise.resolve()
    })

  it('surfaces file-index hits under their own group and opens them', async () => {
    const searched: string[] = []
    mountApi([])
    await render(
      providersWithFiles(async (q) => {
        searched.push(q)
        return {
          hits: [hit({ path: 'C:\\Docs\\repo-note.docx', name: 'repo-note.docx' })],
          total: 1,
          index: { indexed: 1, pending: 0, scanning: false },
        }
      }),
    )
    await actType('repo')
    await flushSearch()
    expect(searched).toEqual(['repo'])
    const groupLabels = [...host.querySelectorAll('.cmd-group-title')].map((el) => el.textContent)
    expect(groupLabels).toContain('Files')
    expect(rowTitles()).toContain('repo-note.docx')
    await clickRow('repo-note.docx')
    expect(openedPaths).toEqual(['C:\\Docs\\repo-note.docx'])
    expect(closed).toBe(1)
  })

  it('a file hit shadows the recent row of the same path', async () => {
    mountApi([recent({ path: 'C:\\Docs\\quarterly.docx', name: 'quarterly.docx' })])
    await render(
      providersWithFiles(async () => ({
        hits: [hit({ path: 'C:\\Docs\\quarterly.docx', name: 'quarterly.docx' })],
        total: 1,
        index: { indexed: 1, pending: 0, scanning: false },
      })),
    )
    await actType('quarterly')
    await flushSearch()
    // one row for the path, under Files; the Ask AI fallback trails
    expect(rowTitles()).toEqual(['quarterly.docx', 'Ask AI: quarterly'])
    const groupLabels = [...host.querySelectorAll('.cmd-group-title')].map((el) => el.textContent)
    expect(groupLabels).toEqual(['Files', 'Ask AI'])
  })

  it('a failing search degrades to no file rows while the rest still answer', async () => {
    mountApi([recent({ path: 'C:\\Docs\\quarterly.docx', name: 'quarterly.docx' })])
    await render(providersWithFiles(() => Promise.reject(new Error('index busy'))))
    await actType('quarterly')
    await flushSearch()
    expect(rowTitles()).toEqual(['quarterly.docx', 'Ask AI: quarterly'])
  })
})

import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * The template center's two channels. The renderer only ever sends a template
 * id; the file name comes from the shipped manifest, and the copy must land in
 * the default save folder under a unique name before the matching app tab
 * opens it. Electron, the docs-main helpers and the tab manager are mocked so
 * only the wiring is under test.
 */

const shared = vi.hoisted(() => ({
  handles: new Map<string, (...args: never[]) => unknown>(),
  recordRecentCalls: [] as string[],
  openedTabs: [] as Array<{ path: string; kind: string }>,
  saveDir: '',
}))

vi.mock('electron', () => ({
  app: { isPackaged: false, getAppPath: () => '/app', getPath: () => '/app/userData' },
  ipcMain: {
    handle: (channel: string, handler: (...args: never[]) => unknown) =>
      shared.handles.set(channel, handler),
  },
}))

// uniquePathIn with its real logic: the copy must never clobber an existing file
vi.mock('../../docs/src/main/docs-main', async () => {
  const { existsSync: exists } = await import('node:fs')
  const { join: joinPath } = await import('node:path')
  return {
    defaultSaveDir: () => shared.saveDir,
    uniquePathIn: (dir: string, fileName: string) => {
      const dot = fileName.lastIndexOf('.')
      const base = dot > 0 ? fileName.slice(0, dot) : fileName
      const ext = dot > 0 ? fileName.slice(dot) : ''
      let candidate = joinPath(dir, fileName)
      for (let i = 2; exists(candidate); i++) candidate = joinPath(dir, `${base}-${i}${ext}`)
      return candidate
    },
    recordRecentFile: (path: string) => void shared.recordRecentCalls.push(path),
  }
})

vi.mock('../src/main/tab-manager', () => ({
  shellTabManager: () => ({
    openDocsTab: (path: string) => void shared.openedTabs.push({ path, kind: 'docs' }),
    openSheetsTab: (path: string) => void shared.openedTabs.push({ path, kind: 'sheets' }),
    openSlidesTab: (path: string) => void shared.openedTabs.push({ path, kind: 'slides' }),
    openMarkdownTab: (path: string) => void shared.openedTabs.push({ path, kind: 'markdown' }),
  }),
}))

import { registerTemplatesIpc } from '../src/main/templates-ipc'

let templatesDir = ''

const invoke = (channel: string, ...args: unknown[]) =>
  shared.handles.get(channel)!({ sender: { id: 1 } } as never, ...args)

const DOCX = 'PK\x03\x04docx-bytes'
const PPTX = 'PK\x03\x04pptx-bytes'

const ENTRIES = [
  {
    id: 'weekly-report',
    title: '周报',
    description: '项目周报',
    kind: 'docx',
    category: 'office',
    file: 'weekly-report.docx',
  },
  {
    id: 'budget',
    title: '团队预算表',
    description: '预算表',
    kind: 'xlsx',
    category: 'finance',
    file: 'budget.xlsx',
  },
  {
    id: 'kickoff',
    title: '项目启动会',
    description: '启动会',
    kind: 'pptx',
    category: 'business',
    file: 'kickoff.pptx',
  },
  {
    id: 'reading-notes',
    title: '读书笔记',
    description: '笔记',
    kind: 'md',
    category: 'education',
    file: 'reading-notes.md',
  },
] as const

function writeFixture(manifest: unknown, files: Record<string, string>): void {
  templatesDir = mkdtempSync(join(tmpdir(), 'tpl-templates-'))
  shared.saveDir = mkdtempSync(join(tmpdir(), 'tpl-save-'))
  writeFileSync(join(templatesDir, 'templates.json'), JSON.stringify(manifest))
  for (const [name, body] of Object.entries(files)) {
    writeFileSync(join(templatesDir, name), body)
  }
}

beforeEach(() => {
  shared.handles.clear()
  shared.recordRecentCalls.length = 0
  shared.openedTabs.length = 0
  writeFixture(ENTRIES, {
    'weekly-report.docx': DOCX,
    'budget.xlsx': 'PK\x03\x04xlsx-bytes',
    'kickoff.pptx': PPTX,
    'reading-notes.md': '# 读书笔记\n',
  })
  registerTemplatesIpc({
    templatesDir: () => templatesDir,
    defaultSaveDir: () => shared.saveDir,
  })
})

afterEach(() => {
  if (templatesDir) rmSync(templatesDir, { recursive: true, force: true })
  if (shared.saveDir) rmSync(shared.saveDir, { recursive: true, force: true })
})

describe('home:templates-list', () => {
  it('lists the manifest entries as TemplateInfo, without the file field', async () => {
    expect(await invoke('home:templates-list')).toEqual(
      ENTRIES.map(({ file: _file, ...info }) => info),
    )
  })

  it('drops entries whose file is not on disk', async () => {
    rmSync(join(templatesDir, 'budget.xlsx'))
    const listed = (await invoke('home:templates-list')) as Array<{ id: string }>
    expect(listed.map((entry) => entry.id)).toEqual(['weekly-report', 'kickoff', 'reading-notes'])
  })

  it('is the empty list when the manifest is missing or corrupt', async () => {
    rmSync(join(templatesDir, 'templates.json'))
    expect(await invoke('home:templates-list')).toEqual([])
    writeFileSync(join(templatesDir, 'templates.json'), 'not json')
    expect(await invoke('home:templates-list')).toEqual([])
  })

  it('drops entries that lost their shape', async () => {
    writeFileSync(
      join(templatesDir, 'templates.json'),
      JSON.stringify([ENTRIES[0], { id: 'broken', kind: 'exe' }]),
    )
    const listed = (await invoke('home:templates-list')) as Array<{ id: string }>
    expect(listed.map((entry) => entry.id)).toEqual(['weekly-report'])
  })
})

describe('home:new-from-template', () => {
  it('copies the template into the default save dir and opens a docs tab', async () => {
    await expect(invoke('home:new-from-template', 'weekly-report')).resolves.toBeUndefined()
    const copied = join(shared.saveDir, '周报.docx')
    expect(existsSync(copied)).toBe(true)
    expect(readFileSync(copied, 'latin1')).toBe(DOCX)
    expect(shared.openedTabs).toEqual([{ path: copied, kind: 'docs' }])
    expect(shared.recordRecentCalls).toEqual([copied])
  })

  it('routes each kind to its own app tab', async () => {
    await invoke('home:new-from-template', 'weekly-report')
    await invoke('home:new-from-template', 'budget')
    await invoke('home:new-from-template', 'kickoff')
    await invoke('home:new-from-template', 'reading-notes')
    expect(shared.openedTabs).toEqual([
      { path: join(shared.saveDir, '周报.docx'), kind: 'docs' },
      { path: join(shared.saveDir, '团队预算表.xlsx'), kind: 'sheets' },
      { path: join(shared.saveDir, '项目启动会.pptx'), kind: 'slides' },
      { path: join(shared.saveDir, '读书笔记.md'), kind: 'markdown' },
    ])
  })

  it('never clobbers: a second copy lands under a unique name', async () => {
    await invoke('home:new-from-template', 'kickoff')
    await invoke('home:new-from-template', 'kickoff')
    expect(readFileSync(join(shared.saveDir, '项目启动会.pptx'), 'latin1')).toBe(PPTX)
    expect(existsSync(join(shared.saveDir, '项目启动会-2.pptx'))).toBe(true)
    expect(readFileSync(join(shared.saveDir, '项目启动会-2.pptx'), 'latin1')).toBe(PPTX)
  })

  it('refuses an unknown id, a file name and a path traversal', async () => {
    for (const id of ['nope', 'weekly-report.docx', '../outside']) {
      await expect(invoke('home:new-from-template', id)).rejects.toThrow(/no template/)
    }
    expect(existsSync(join(shared.saveDir, '周报.docx'))).toBe(false)
  })

  it('refuses a non-string id', async () => {
    await expect(invoke('home:new-from-template', '')).rejects.toThrow(/template id/)
    await expect(invoke('home:new-from-template', { id: 'weekly-report' })).rejects.toThrow(
      /template id/,
    )
  })

  it('refuses a listed template whose file went missing', async () => {
    rmSync(join(templatesDir, 'weekly-report.docx'))
    await expect(invoke('home:new-from-template', 'weekly-report')).rejects.toThrow(
      /missing its file/,
    )
  })

  it('keeps the copy even when no shell window is left to open a tab in', async () => {
    // a detached editor outlived the shell: the file still lands, the call still succeeds
    registerTemplatesIpc({
      templatesDir: () => templatesDir,
      defaultSaveDir: () => shared.saveDir,
      openTemplate: () => undefined,
    })
    await expect(invoke('home:new-from-template', 'weekly-report')).resolves.toBeUndefined()
    expect(existsSync(join(shared.saveDir, '周报.docx'))).toBe(true)
    expect(shared.openedTabs).toEqual([])
  })
})

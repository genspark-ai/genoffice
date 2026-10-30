/**
 * The desktop's item model and provider registry (src/renderer/src/desktop-items.ts):
 * what each source contributes, the cross-source dedupe, the alias open
 * routing (suite tab vs OS default), the drop-path resolution and the
 * widgets' reserve band. Pure functions — no bridge, no DOM.
 */
import { describe, expect, it, vi } from 'vitest'
import type { DesktopPinInfo, FolderRoot, RecentEntry } from '../src/shared/home-api'
import {
  defaultDesktopProviders,
  droppedFilePaths,
  isOpenableDocument,
  mergeDesktopItems,
  widgetReservedCellKeys,
  WIDGET_RESERVE_CELLS,
  type DesktopItemContext,
  type DesktopItemProvider,
} from '../src/renderer/src/desktop-items'
import { desktopStrings } from '../src/renderer/src/desktop-strings'

const strings = desktopStrings('en')

const entry = (path: string, extra: Partial<RecentEntry> = {}): RecentEntry => ({
  path,
  name: path.split('/').pop() ?? path,
  ext: (path.slice(path.lastIndexOf('.') + 1) || '').toLowerCase(),
  mtimeMs: 1,
  sizeBytes: 10,
  starred: false,
  ...extra,
})

const root = (path: string, name: string): FolderRoot => ({
  path,
  name,
  usable: true,
  readable: true,
  removable: true,
})

const pin = (over: Partial<DesktopPinInfo> = {}): DesktopPinInfo => ({
  id: 'detached:22',
  kind: 'sheets',
  title: 'budget.xlsx',
  filePath: '/docs/budget.xlsx',
  bounds: { x: 600, y: 400, width: 480, height: 360 },
  ...over,
})

const ctx = (over: Partial<DesktopItemContext> = {}): DesktopItemContext => ({
  strings,
  roots: [root('/docs/Work', 'Work')],
  recents: [entry('/docs/alpha.docx'), entry('/docs/beta.xlsx')],
  aliases: [entry('/docs/report.docx')],
  aliasPaths: new Set(['/docs/report.docx']),
  pins: [],
  api: {},
  onOpenAsk: () => {},
  onOpenTemplates: () => {},
  onOpenSettings: () => {},
  onOpenFolder: () => {},
  ...over,
})

const titles = (items: ReturnType<typeof mergeDesktopItems>): string[] => items.map((i) => i.title)

describe('defaultDesktopProviders', () => {
  it('registers the five sources in display order', () => {
    expect(defaultDesktopProviders().map((provider) => provider.id)).toEqual([
      'apps',
      'folder-roots',
      'recents',
      'desktop-files',
      'widgets',
    ])
  })
})

describe('mergeDesktopItems', () => {
  it('folds every source in registry order: apps, roots, recents, aliases, widgets', () => {
    const items = mergeDesktopItems(defaultDesktopProviders(), ctx())
    expect(titles(items)).toEqual([
      'Docs',
      'Sheets',
      'Slides',
      'PDF',
      'Markdown',
      'Web Page',
      'AI Chat',
      'Templates',
      'Plugins',
      'Settings',
      'Work',
      'alpha.docx',
      'beta.xlsx',
      'report.docx',
    ])
    // ids carry the source: the alias keeps its own id namespace
    expect(items.map((i) => i.id)).toContain('alias:/docs/report.docx')
    expect(items.map((i) => i.id)).toContain('file:/docs/alpha.docx')
    expect(items.every((i) => i.kind !== 'widget')).toBe(true)
  })

  it('an alias replaces the recent and the folder root of the same path', () => {
    const items = mergeDesktopItems(
      defaultDesktopProviders(),
      ctx({
        recents: [entry('/docs/alpha.docx'), entry('/docs/report.docx')],
        roots: [root('/docs/Work', 'Work'), root('/docs/report.docx', 'report.docx')],
      }),
    )
    // report.docx appears once — as the alias, not as the recent or the root
    expect(items.filter((i) => i.path === '/docs/report.docx')).toHaveLength(1)
    expect(items.find((i) => i.path === '/docs/report.docx')?.alias).toBe(true)
    expect(items.find((i) => i.path === '/docs/report.docx')?.kind).toBe('file')
  })

  it('marks only the dragged-in items as aliases; every non-widget item drags', () => {
    const items = mergeDesktopItems(defaultDesktopProviders(), ctx())
    for (const item of items) {
      expect(item.alias).toBe(item.id.startsWith('alias:'))
      expect(item.draggable).toBe(true)
    }
  })

  it('widget items carry their pin and do not join the grid set', () => {
    const items = mergeDesktopItems(defaultDesktopProviders(), ctx({ pins: [pin()] }))
    const widget = items.find((i) => i.kind === 'widget')
    expect(widget).toMatchObject({
      id: 'detached:22',
      title: 'budget.xlsx',
      path: '/docs/budget.xlsx',
      draggable: false,
      alias: false,
    })
  })

  it('one icon per path across providers, whatever a source repeats', () => {
    const providers: DesktopItemProvider[] = [
      {
        id: 'recents',
        getItems: () => [
          {
            id: 'file:/a.docx',
            kind: 'file',
            title: 'a.docx',
            tile: null,
            ext: 'docx',
            path: '/a.docx',
            alias: false,
            draggable: true,
            open: () => {},
          },
          {
            id: 'file:/a.docx:2',
            kind: 'file',
            title: 'a.docx again',
            tile: null,
            ext: 'docx',
            path: '/a.docx',
            alias: false,
            draggable: true,
            open: () => {},
          },
        ],
      },
    ]
    expect(mergeDesktopItems(providers, ctx())).toHaveLength(1)
  })
})

describe('alias opens', () => {
  it('routes a known document through openPath, the OS default otherwise', () => {
    const api = {
      openPath: vi.fn(),
      desktopOpenExternal: vi.fn(),
    }
    const aliases = [
      entry('/docs/report.docx'),
      entry('/docs/photo.png'),
      entry('/docs/Projects', { directory: true }),
    ]
    const items = mergeDesktopItems(
      defaultDesktopProviders(),
      ctx({ aliases, api, aliasPaths: new Set(aliases.map((a) => a.path)) }),
    )
    items.filter((i) => i.alias).forEach((item) => item.open())
    expect(api.openPath).toHaveBeenCalledWith('/docs/report.docx')
    expect(api.openPath).toHaveBeenCalledTimes(1)
    expect(api.desktopOpenExternal).toHaveBeenCalledWith('/docs/photo.png')
    // a folder is the OS's to open (in-app browsing is the folder-roots surface)
    expect(api.desktopOpenExternal).toHaveBeenCalledWith('/docs/Projects')
    expect(api.desktopOpenExternal).toHaveBeenCalledTimes(2)
  })

  it('a folder alias renders as a folder, a file alias as a file', () => {
    const aliases = [entry('/docs/report.docx'), entry('/docs/Projects', { directory: true })]
    const items = mergeDesktopItems(
      defaultDesktopProviders(),
      ctx({ aliases, aliasPaths: new Set(aliases.map((a) => a.path)) }),
    )
    expect(items.find((i) => i.path === '/docs/report.docx')?.kind).toBe('file')
    expect(items.find((i) => i.path === '/docs/Projects')?.kind).toBe('folder')
  })

  it('keeps the extension knowledge the drop bridge defines', () => {
    for (const path of ['/a.docx', '/b.XLSX', '/c.md', '/d.markdown', '/e.html', '/f.ots']) {
      expect(isOpenableDocument(path), path).toBe(true)
    }
    for (const path of ['/a.png', '/b.zip', '/c.exe', '/d', '/e.MOV']) {
      expect(isOpenableDocument(path), path).toBe(false)
    }
  })
})

describe('droppedFilePaths', () => {
  const file = (name: string): File => new File(['x'], name)

  it('resolves each dropped file through the bridge, skipping failures and dupes', () => {
    const files = [file('a.docx'), file('b.xlsx'), file('c.docx')]
    const paths = droppedFilePaths(files, (f) =>
      f.name === 'b.xlsx' ? undefined : `/docs/${f.name}`,
    )
    expect(paths).toEqual(['/docs/a.docx', '/docs/c.docx'])
  })

  it('a resolver that throws skips only that file', () => {
    const paths = droppedFilePaths([file('a.docx'), file('bad')], (f) => {
      if (f.name === 'bad') throw new Error('sandboxed entry')
      return `/docs/${f.name}`
    })
    expect(paths).toEqual(['/docs/a.docx'])
  })

  it('caps one drop at 20 aliases and drops pathologically long ones', () => {
    const many = Array.from({ length: 30 }, (_v, i) => file(`f${i}.docx`))
    expect(droppedFilePaths(many, (f) => `/docs/${f.name}`)).toHaveLength(20)
    expect(droppedFilePaths([file('a.docx')], () => `${'/'.repeat(5000)}a.docx`)).toEqual([])
  })
})

describe('widgetReservedCellKeys', () => {
  it('reserves nothing without widgets', () => {
    expect(widgetReservedCellKeys(0, 7, 12).size).toBe(0)
    expect(widgetReservedCellKeys(-1, 7, 12).size).toBe(0)
  })

  it('reserves the bottom-right corner: last column upward, then the next one in', () => {
    // one widget = 12 cells: the last two full columns of a 6-row grid
    const keys = widgetReservedCellKeys(1, 6, 5)
    expect([...keys]).toEqual([
      '4:5',
      '4:4',
      '4:3',
      '4:2',
      '4:1',
      '4:0',
      '3:5',
      '3:4',
      '3:3',
      '3:2',
      '3:1',
      '3:0',
    ])
    // a second widget continues into the third-to-last column
    const two = widgetReservedCellKeys(2, 6, 5)
    expect(two.size).toBe(24)
    expect([...two]).toContain('2:5')
    expect([...two]).toContain('2:0')
    // a partial column happens when the count is not a whole column multiple
    const partial = widgetReservedCellKeys(1, 5, 6)
    expect([...partial].slice(-2)).toEqual(['3:4', '3:3'])
  })

  it('scales with the widget count and never exceeds the grid', () => {
    expect(widgetReservedCellKeys(1, 7, 12).size).toBe(WIDGET_RESERVE_CELLS)
    expect(widgetReservedCellKeys(4, 7, 12).size).toBe(4 * WIDGET_RESERVE_CELLS)
    // a grid smaller than the band is fully reserved, never more
    expect(widgetReservedCellKeys(1, 2, 3).size).toBe(6)
    expect(widgetReservedCellKeys(99, 7, 12).size).toBe(84)
    // degenerate grids answer empty rather than looping
    expect(widgetReservedCellKeys(1, 0, 3).size).toBe(0)
    expect(widgetReservedCellKeys(1, 3, 0).size).toBe(0)
  })
})

// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { defaultAiSettings } from '@genoffice/ai-provider'
import { Home } from '../src/renderer/src/Home'
import { LocaleProvider } from '../src/renderer/src/locale'
import type { RecentEntry } from '../src/shared/home-api'

/**
 * The Home shell around the desktop view: the Desktop nav item sits on top of
 * the sidebar, the desktop is the view Home opens by default, and the last
 * top-level nav the person chose is what the next launch opens.
 */

let root: Root
let container: HTMLDivElement
const files: RecentEntry[] = [
  {
    path: '/fixtures/brief.docx',
    name: 'brief.docx',
    ext: 'docx',
    mtimeMs: 1,
    sizeBytes: 10,
    starred: true,
  },
  {
    path: '/fixtures/budget.xlsx',
    name: 'budget.xlsx',
    ext: 'xlsx',
    mtimeMs: 2,
    sizeBytes: 20,
    starred: false,
  },
]
const page = () =>
  vi.fn(async (query: { ext?: string; limit?: number }) => ({
    entries: query.limit === 0 ? [] : files,
    total: files.length,
    totalAll: files.length,
  }))

beforeEach(async () => {
  HTMLElement.prototype.scrollTo = vi.fn()
  localStorage.clear()
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  vi.stubGlobal(
    'fetch',
    vi.fn(() => {
      throw new Error('No network in desktop navigation tests')
    }),
  )
  window.aiOffice = {
    folderRoots: async () => [],
    onFolderChanged: () => () => {},
    recents: page(),
    starred: page(),
    starredGroups: async () => [],
    accountStatus: async () => ({ loggedIn: false }),
    onOpenSettings: () => () => {},
    onAccountLogin: () => () => {},
    kbList: async () => [],
    aiGetSettings: async () => defaultAiSettings(),
    aiGskStatus: async () => ({ loggedIn: false }),
    onAiSettingsChanged: () => () => {},
    desktopLayout: async () => ({ positions: {} }),
    desktopLayoutSave: vi.fn(async () => undefined),
    onRecentsChanged: () => () => {},
  } as unknown as Window['aiOffice']
  container = document.createElement('div')
  document.body.append(container)
  root = createRoot(container)
  await act(async () =>
    root.render(createElement(LocaleProvider, { initial: 'en', children: createElement(Home) })),
  )
})

afterEach(async () => {
  await act(async () => root.unmount())
  container.remove()
  localStorage.clear()
  vi.unstubAllGlobals()
})

function nav(label: string): HTMLButtonElement {
  return [...container.querySelectorAll<HTMLButtonElement>('.sidebar-nav button')].find(
    (button) => button.querySelector('.nav-label')?.textContent === label,
  )!
}

async function click(button: HTMLButtonElement) {
  expect(button).toBeTruthy()
  await act(async () => button.click())
}

function selected(): string[] {
  return [...container.querySelectorAll('.sidebar-nav .active .nav-label')].map(
    (label) => label.textContent!,
  )
}

describe('Home opens on the desktop', () => {
  it('the Desktop nav item sits first in the sidebar and is active on launch', () => {
    const labels = [...container.querySelectorAll('.sidebar-nav .nav-label')].map(
      (label) => label.textContent,
    )
    expect(labels[0]).toBe('Desktop')
    expect(labels.slice(1)).toEqual(['Recent', 'Starred', 'Ask AI'])
    expect(selected()).toEqual(['Desktop'])
    expect(nav('Desktop').getAttribute('aria-current')).toBe('page')
    // the content area is the desktop, not the recents list
    expect(container.querySelector('.desktop-surface')).toBeTruthy()
    expect(container.querySelector('.desktop-icons')).toBeTruthy()
    expect(container.querySelector('.recent-table')).toBeNull()
  })

  it('the wallpaper backdrop is lifted to the Home root, glass on the sidebar', () => {
    expect(container.querySelector('.home')!.className).toContain('desktop-mode')
    // the host Home mounts at its root carries the desktop's backdrop (aurora
    // behind the default look), and the sidebar wears its glass pane
    expect(container.querySelector('.home-desktop-backdrop')).toBeTruthy()
    expect(container.querySelector('.sidebar-glass')).toBeTruthy()
    expect(container.querySelector('.home-desktop-backdrop .desktop-backdrop')).toBeTruthy()
    expect(container.querySelector('.home-desktop-backdrop .desktop-aurora-a')).toBeTruthy()
    // the surface itself carries no backdrop of its own anymore
    expect(container.querySelector('.desktop-surface .desktop-backdrop')).toBeNull()
  })

  it('leaving the desktop returns the window to the plain list-view look', async () => {
    await click(nav('Recent'))
    expect(container.querySelector('.home')!.className).not.toContain('desktop-mode')
    // the lifted wallpaper, its host and the sidebar glass all unmount
    expect(container.querySelector('.home-desktop-backdrop')).toBeNull()
    expect(container.querySelector('.sidebar-glass')).toBeNull()
    expect(container.querySelector('.desktop-backdrop')).toBeNull()
    expect(container.querySelector('.desktop-surface')).toBeNull()
    expect(container.querySelector('.recent-table')).toBeTruthy()
    await click(nav('Desktop'))
    expect(container.querySelector('.home-desktop-backdrop .desktop-backdrop')).toBeTruthy()
    expect(container.querySelector('.sidebar-glass')).toBeTruthy()
  })

  it('the desktop shows the icon set: apps and the newest recents', () => {
    const labels = [...container.querySelectorAll('.desktop-icon-label')].map((l) => l.textContent)
    expect(labels).toEqual(
      expect.arrayContaining([
        'Docs',
        'Sheets',
        'Slides',
        'AI Chat',
        'Templates',
        'Settings',
        'brief.docx',
        'budget.xlsx',
      ]),
    )
  })

  it('switching away and back keeps the desktop a first-class view', async () => {
    await click(nav('Recent'))
    expect(selected()).toEqual(['Recent'])
    expect(container.querySelector('.desktop-surface')).toBeNull()
    expect(container.querySelector('.recent-table')).toBeTruthy()
    await click(nav('Desktop'))
    expect(selected()).toEqual(['Desktop'])
    expect(container.querySelector('.desktop-surface')).toBeTruthy()
    await click(nav('Ask AI'))
    expect(selected()).toEqual(['Ask AI'])
    expect(container.querySelector('.kb-ask')).toBeTruthy()
    await click(nav('Desktop'))
    expect(container.querySelector('.kb-ask')).toBeNull()
    expect(container.querySelector('.desktop-surface')).toBeTruthy()
  })
})

describe('the launch view persists', () => {
  it('the last top-level nav wins on the next mount', async () => {
    await click(nav('Starred'))
    expect(localStorage.getItem('home.view')).toBe('starred')
    await act(async () => root.unmount())
    container.remove()
    container = document.createElement('div')
    document.body.append(container)
    root = createRoot(container)
    await act(async () =>
      root.render(createElement(LocaleProvider, { initial: 'en', children: createElement(Home) })),
    )
    expect(selected()).toEqual(['Starred'])
    expect(container.querySelector('.recent-table')).toBeTruthy()
    expect(container.querySelector('.desktop-surface')).toBeNull()
  })

  it('the default with no stored preference (or a corrupt one) is the desktop', async () => {
    localStorage.setItem('home.view', 'nonsense')
    await act(async () => root.unmount())
    container.remove()
    container = document.createElement('div')
    document.body.append(container)
    root = createRoot(container)
    await act(async () =>
      root.render(createElement(LocaleProvider, { initial: 'en', children: createElement(Home) })),
    )
    expect(selected()).toEqual(['Desktop'])
    expect(container.querySelector('.desktop-surface')).toBeTruthy()
  })

  it('a folder picked in the tree takes over the content area from the desktop', async () => {
    window.aiOffice = {
      ...window.aiOffice,
      folderRoots: async () => [
        { path: '/docs', name: 'Docs', usable: true, readable: true, removable: false },
      ],
      listFolder: async () => ({ dir: '/docs', folders: [], files }),
    } as unknown as Window['aiOffice']
    await act(async () => root.unmount())
    container.remove()
    container = document.createElement('div')
    document.body.append(container)
    root = createRoot(container)
    await act(async () =>
      root.render(createElement(LocaleProvider, { initial: 'en', children: createElement(Home) })),
    )
    await act(async () => {
      const row = container.querySelector<HTMLElement>('.tree-row')
      row?.click()
    })
    expect(container.querySelector('.desktop-surface')).toBeNull()
    expect(container.querySelector('.recent-table')).toBeTruthy()
  })
})

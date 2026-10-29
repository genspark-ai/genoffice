/**
 * @vitest-environment jsdom
 */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import type { Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { HomeApi } from '../src/shared/home-api'
import type { InstalledPlugin } from '../src/shared/found-skill'
import { PluginsPane } from '../src/renderer/src/PluginsPane'
import { SettingsModal, type SettingsModalProps } from '../src/renderer/src/SettingsModal'
import { LocaleProvider } from '../src/renderer/src/locale'
import { en as enStrings } from '../src/renderer/src/strings/en'

/** the real English text, so assertions read like what a user sees */
const en = enStrings as Record<string, string>
const t = ((key: string, params?: Record<string, string>) =>
  (en[key] ?? key).replace(/\{(\w+)\}/g, (_, name) => params?.[name] ?? `{${name}}`)) as never

/**
 * The pane that lists the plugins installed into GenOffice and installs more
 * from a folder or a zip. The only thing that copies anything is pressing one
 * of those two buttons; the list itself is read-only, and every action is
 * followed by a relist rather than trusted.

 * The mock bridge below behaves like a tiny main process: it keeps the folder
 * state, applies writes to it, and answers listings from it.
 */

const actEnvironment = globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
actEnvironment.IS_REACT_ACT_ENVIRONMENT = true

let host: HTMLDivElement
let root: Root

const plugin = (over: Partial<InstalledPlugin>): InstalledPlugin => ({
  name: 'pdf-toolkit',
  version: '1.2.0',
  description: 'Assemble and fix PDF documents',
  author: 'Genspark Labs',
  enabled: true,
  components: {
    skills: ['pdf-toolkit:assemble', 'pdf-toolkit:fix'],
    commands: ['pdf-toolkit:stamp'],
    mcpServers: ['pdf-toolkit:pdf-server'],
  },
  format: 'claude',
  ...over,
})

/** the folder state the mock main process answers listings from */
let store: InstalledPlugin[]
/** what the two install dialogs hand back (null = the user cancelled) */
let zipAnswer: InstalledPlugin | null
let folderAnswer: InstalledPlugin | null
const imports: string[] = []
const toggles: Array<[string, boolean]> = []
const removes: string[] = []
let listCalls = 0

beforeEach(() => {
  host = document.createElement('div')
  document.body.append(host)
  root = createRoot(host)
  store = []
  zipAnswer = null
  folderAnswer = null
  imports.length = 0
  toggles.length = 0
  removes.length = 0
  listCalls = 0
})

afterEach(() => {
  act(() => root.unmount())
  host.remove()
  vi.restoreAllMocks()
})

function mountApi(list: InstalledPlugin[]): void {
  store = list
  window.aiOffice = pluginBridge() as unknown as HomeApi
}

/** the plugins half of window.aiOffice, over the folder state above */
function pluginBridge(): Record<string, unknown> {
  return {
    pluginsList: async () => {
      listCalls += 1
      return store
    },
    pluginsImportZip: async () => {
      imports.push('zip')
      if (zipAnswer) store = [...store, zipAnswer]
      return zipAnswer
    },
    pluginsImportFolder: async () => {
      imports.push('folder')
      if (folderAnswer) store = [...store, folderAnswer]
      return folderAnswer
    },
    pluginsSetEnabled: async (name: string, enabled: boolean) => {
      toggles.push([name, enabled])
      store = store.map((p) => (p.name === name ? { ...p, enabled } : p))
    },
    pluginsRemove: async (name: string) => {
      removes.push(name)
      store = store.filter((p) => p.name !== name)
    },
  }
}

/** the settings the modal itself asks of window.aiOffice while it opens */
function baseHomeApi(): Record<string, unknown> {
  return {
    getTheme: async () => 'system',
    getDefaultSaveDir: async () => '',
    getAnalyticsEnabled: async () => true,
    getAutoSaveDefault: async () => ({ on: false, updatedAt: 0 }),
    getAiPanelPrefs: async () => ({ fontSize: 'default', spellcheck: true }),
    getUpdateChannel: async () => 'stable',
    getAppVersion: async () => '1.0.0',
    githubStars: async () => null,
  }
}

async function render(): Promise<void> {
  await act(async () => {
    root.render(createElement(PluginsPane, { t }))
  })
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
  })
}

async function click(el: Element | null | undefined): Promise<void> {
  expect(el).toBeTruthy()
  await act(async () => {
    ;(el as HTMLElement).click()
    await Promise.resolve()
  })
  await act(async () => {
    await Promise.resolve()
    await Promise.resolve()
  })
}

const buttonWithText = (text: string): HTMLButtonElement | undefined =>
  [...host.querySelectorAll('button')].find((b) => b.textContent?.trim() === text)

/** the row that names this plugin, and its enabled switch */
const row = (name: string): Element | undefined =>
  [...host.querySelectorAll('.set-plug-row')].find((r) =>
    r.querySelector('.set-field-label')?.textContent?.includes(name),
  )
const switchOf = (name: string): HTMLButtonElement | null =>
  row(name)?.querySelector<HTMLButtonElement>('button[role="switch"]') ?? null

describe('PluginsPane', () => {
  it('renders nothing until the listing answers', async () => {
    mountApi([plugin({})])
    let resolve: (v: InstalledPlugin[]) => void = () => {}
    ;(window.aiOffice as HomeApi).pluginsList = () =>
      new Promise<InstalledPlugin[]>((r) => (resolve = r))
    await act(async () => {
      root.render(createElement(PluginsPane, { t }))
    })
    // an empty section that fills in a moment later reads as "none installed"
    expect(host.textContent).toBe('')

    await act(async () => {
      resolve([plugin({})])
      await Promise.resolve()
    })
    expect(host.textContent).toContain('pdf-toolkit')
  })

  it('says so when nothing is installed, with the install buttons', async () => {
    mountApi([])
    await render()
    expect(host.textContent).toContain('No plugins installed yet.')
    expect(buttonWithText('Install plugin package (zip)…')).toBeTruthy()
    expect(buttonWithText('Install plugin folder…')).toBeTruthy()
  })

  it('lists a plugin with its byline, component chips, MCP note and format badge', async () => {
    mountApi([plugin({})])
    await render()
    const text = host.textContent ?? ''
    expect(text).toContain('pdf-toolkit')
    expect(text).toContain('1.2.0 · Genspark Labs')
    expect(text).toContain('Assemble and fix PDF documents')
    expect(text).toContain('2 skills')
    expect(text).toContain('1 command')
    expect(text).toContain('Declares MCP servers; they are listed here but not run.')
    // the format badge is its own element: the lead text names Claude Code too
    expect(row('pdf-toolkit')?.querySelector('.set-plug-badge')?.textContent).toBe('Claude Code')
    expect(switchOf('pdf-toolkit')?.getAttribute('aria-checked')).toBe('true')
  })

  it('skips the badge, chips and MCP note when there is nothing to show for them', async () => {
    mountApi([
      plugin({
        name: 'plain-helper',
        components: { skills: [], commands: [], mcpServers: [] },
        format: 'plain',
      }),
    ])
    await render()
    expect(host.querySelector('.set-plug-badge')).toBeNull()
    expect(host.querySelector('.set-plug-chip')).toBeNull()
    expect(host.textContent).not.toContain('Declares MCP servers')
  })

  it('toggles through the bridge and takes the relist as the answer', async () => {
    mountApi([plugin({})])
    await render()
    expect(listCalls).toBe(1)
    await click(switchOf('pdf-toolkit'))
    expect(toggles).toEqual([['pdf-toolkit', false]])
    // the write was followed by a relist, and the listing answered with the
    // write applied: the row shows the new state
    expect(listCalls).toBe(2)
    expect(switchOf('pdf-toolkit')?.getAttribute('aria-checked')).toBe('false')
  })

  it('keeps the row on what the listing says when the write changes nothing', async () => {
    mountApi([plugin({})])
    // a write the folder never accepted: the listing keeps answering "enabled"
    ;(window.aiOffice as HomeApi).pluginsSetEnabled = async () => undefined
    await render()
    await click(switchOf('pdf-toolkit'))
    // the relist ran, and the row stayed on its answer rather than the click
    expect(listCalls).toBe(2)
    expect(switchOf('pdf-toolkit')?.getAttribute('aria-checked')).toBe('true')
  })

  it('asks before removing, then removes and relists', async () => {
    mountApi([plugin({})])
    const confirm = vi.spyOn(window, 'confirm').mockReturnValue(false)
    await render()
    await click(buttonWithText('Remove'))
    expect(confirm).toHaveBeenCalledWith(
      'Remove the plugin "pdf-toolkit"? Its folder is deleted from GenOffice.',
    )
    // declined: nothing left the folder
    expect(removes).toEqual([])
    confirm.mockReturnValue(true)
    await click(buttonWithText('Remove'))
    expect(removes).toEqual(['pdf-toolkit'])
    expect(host.textContent).toContain('Removed pdf-toolkit')
    // the relist is empty now: the row is gone and the empty state is back
    expect(switchOf('pdf-toolkit')).toBeNull()
    expect(host.textContent).toContain('No plugins installed yet.')
  })

  it('says nothing when the zip dialog is cancelled', async () => {
    mountApi([plugin({})])
    await render()
    await click(buttonWithText('Install plugin package (zip)…'))
    expect(imports).toEqual(['zip'])
    expect(host.textContent).not.toContain('Installed')
    expect(host.textContent).not.toContain('Install failed')
  })

  it('names the imported plugin in the notice and shows it in the list', async () => {
    mountApi([])
    folderAnswer = plugin({
      name: 'fresh-plugin',
      components: { skills: ['fresh-plugin:go'], commands: [], mcpServers: [] },
    })
    await render()
    expect(host.textContent).toContain('No plugins installed yet.')
    await click(buttonWithText('Install plugin folder…'))
    expect(imports).toEqual(['folder'])
    expect(host.textContent).toContain('Installed fresh-plugin')
    expect(host.textContent).toContain('fresh-plugin')
    expect(host.textContent).not.toContain('No plugins installed yet.')
  })

  it('shows the reason when an install is refused', async () => {
    mountApi([plugin({})])
    ;(window.aiOffice as HomeApi).pluginsImportZip = async () => {
      throw new Error('the archive contains a symlink, and installs refuse those')
    }
    await render()
    await click(buttonWithText('Install plugin package (zip)…'))
    expect(host.textContent).toContain('Install failed: the archive contains a symlink')
  })

  it('survives a preload that predates the feature', async () => {
    // an app update can leave the old bridge in place for one launch
    window.aiOffice = {} as HomeApi
    await render()
    expect(host.textContent).toBe('')
  })

  it('opens as its own section of the settings modal, right after skills', async () => {
    store = [plugin({})]
    window.aiOffice = { ...baseHomeApi(), ...pluginBridge() } as unknown as HomeApi
    await act(async () => {
      root.render(
        createElement(
          LocaleProvider,
          { initial: 'en' },
          createElement(SettingsModal, {
            status: null,
            loggingOut: false,
            loginWaiting: false,
            loginUrl: null,
            urlCopied: false,
            onOpenLoginUrl: vi.fn(),
            onCopyLoginUrl: vi.fn(),
            onClose: vi.fn(),
            onLogin: vi.fn(),
            onLogout: vi.fn(),
          } satisfies SettingsModalProps),
        ),
      )
      await Promise.resolve()
    })
    // the nav lists the section, between skills and about
    const nav = [...host.querySelectorAll<HTMLButtonElement>('.set-nav-item')]
    const labels = nav.map((b) => b.textContent)
    expect(labels.indexOf('Skills')).toBeGreaterThan(-1)
    expect(labels.indexOf('Plugins')).toBe(labels.indexOf('Skills') + 1)
    expect(labels.indexOf('Plugins')).toBeLessThan(labels.indexOf('About'))
    await click(buttonWithText('Plugins'))
    // the pane mounted through the section, listing from the real bridge
    expect(host.textContent).toContain('pdf-toolkit')
    expect(buttonWithText('Install plugin package (zip)…')).toBeTruthy()
  })
})

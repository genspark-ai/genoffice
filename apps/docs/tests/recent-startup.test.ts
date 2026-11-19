// @vitest-environment node
import { readFileSync } from 'node:fs'
import { basename } from 'node:path'
import { createContext, runInContext } from 'node:vm'
import ts from 'typescript'
import { expect, it, vi } from 'vitest'
import { COMMANDS } from '../src/shared/commands'
import {
  bindingId,
  keepsAccelerator,
  menuAccelerator,
  type KeyOverrides,
} from '../src/shared/keybindings'
import { DOCS_MENU } from '../src/shared/menu-layout'

const source = ts.createSourceFile(
  'docs-main.ts',
  readFileSync(new URL('../src/main/docs-main.ts', import.meta.url), 'utf8'),
  ts.ScriptTarget.Latest,
  true,
)

function fixture(disabled: string[] = [], overrides: KeyOverrides = {}) {
  const paths = ['/offline/share.docx', '/local/document.docx']
  const existsSync = vi.fn(() => {
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 50)
    return true
  })
  const installMenu = vi.fn((_menu: unknown) => {})
  let recentHandler: (() => unknown) | undefined
  // Execute the production entry points without booting Electron or opening a window.
  const context = createContext({
    exports: {},
    process: { platform: 'linux' },
    docsMenuGate: null,
    isDev: false,
    readJson: () => paths,
    RECENT_PATH: () => '/synthetic/recent.json',
    existsSync,
    basename,
    COMMANDS,
    DOCS_MENU,
    tm: (key: string) => key,
    getUiLang: () => 'en',
    appMenuLabels: () => ({}),
    appMenuTemplate: () => null,
    settingsMenuItem: () => ({}),
    windowMenuTemplate: () => ({}),
    checkUpdatesMenuItem: () => ({}),
    aboutMenuItem: () => ({}),
    app: { name: 'Synthetic' },
    extraFileMenuItems: [],
    shellHooks: null,
    activeCommandState: () => ({ checked: [], disabled }),
    COMMAND_ITEM_ID: 'docs-cmd:',
    bindingId,
    menuAccelerator,
    keepsAccelerator,
    userKeyOverrides: () => overrides,
    OWN_MENU_ITEM: new Set(['bold', 'file-properties']),
    Menu: { buildFromTemplate: (template: unknown) => template, setApplicationMenu: installMenu },
    ipcMain: {
      handle: (_channel: string, handler: () => unknown) => {
        recentHandler = handler
      },
    },
  })
  const execute = (code: string) =>
    runInContext(
      ts.transpileModule(code, {
        compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 },
      }).outputText,
      context,
    )
  return { paths, existsSync, installMenu, execute, recent: () => recentHandler!() }
}

it('builds the startup Docs menu without probing recent paths', () => {
  const f = fixture()
  const node = source.statements.find(
    (n) => ts.isFunctionDeclaration(n) && n.name?.text === 'buildDocsMenu',
  )!
  f.execute(node.getText(source))
  f.execute('buildDocsMenu()')
  expect(f.existsSync).not.toHaveBeenCalled()
  expect(f.installMenu).toHaveBeenCalledOnce()
})

it('greys out the host-built items of a command reported unavailable', () => {
  const f = fixture(['name-with-ai', 'zoom-set', 'save'])
  const node = source.statements.find(
    (n) => ts.isFunctionDeclaration(n) && n.name?.text === 'buildDocsMenu',
  )!
  f.execute(node.getText(source))
  f.execute('buildDocsMenu()')
  type Item = { id?: string; enabled?: boolean; submenu?: Item[] }
  const flat = (items: Item[]): Item[] => items.flatMap((i) => [i, ...flat(i.submenu ?? [])])
  const items = flat(f.installMenu.mock.calls[0][0] as Item[])
  const of = (id: string) => items.filter((i) => i.id === `docs-cmd:${id}`)
  expect(of('name-with-ai').map((i) => i.enabled)).toEqual([false])
  expect(of('zoom-set')).toHaveLength(9)
  expect(of('zoom-set').every((i) => i.enabled === false)).toBe(true)
  expect(of('save').map((i) => i.enabled)).toEqual([false])
  expect(of('open').every((i) => i.enabled !== false)).toBe(true)
})

it('shows a customized chord on its menu item and drops one that moved away', () => {
  const f = fixture([], { bold: [], 'file-properties': ['Ctrl+Shift+F7'], open: ['Ctrl+O'] })
  const node = source.statements.find(
    (n) => ts.isFunctionDeclaration(n) && n.name?.text === 'buildDocsMenu',
  )!
  f.execute(node.getText(source))
  f.execute('buildDocsMenu()')
  type Item = { id?: string; label?: string; accelerator?: string; submenu?: Item[] }
  const flat = (items: Item[]): Item[] => items.flatMap((i) => [i, ...flat(i.submenu ?? [])])
  const items = flat(f.installMenu.mock.calls[0][0] as Item[])
  const chords = (id: string) =>
    items.filter((i) => i.id === `docs-cmd:${id}`).map((i) => i.accelerator)
  expect(chords('file-properties')).toEqual(['Ctrl+Shift+F7'])
  expect(chords('bold')).toEqual([undefined])
  expect(chords('italic')).toEqual(['CmdOrCtrl+I'])
  // the host-built item that carries Open's second chord gives it up with Open
  expect(items.find((i) => i.label === 'menuRecentMore')?.accelerator).toBe(undefined)
  expect(chords('open')).toEqual(['Ctrl+O'])
})

it('returns saved recent paths without checking offline mounts at the legacy startup IPC', () => {
  const f = fixture()
  let registration: ts.CallExpression | undefined
  function visit(node: ts.Node) {
    if (
      ts.isCallExpression(node) &&
      node.expression.getText(source) === 'ipcMain.handle' &&
      node.arguments[0]?.getText(source) === "'docs:recent'"
    )
      registration = node
    ts.forEachChild(node, visit)
  }
  visit(source)
  expect(registration).toBeDefined()
  f.execute(registration!.getText(source))
  expect(f.recent()).toEqual(f.paths)
  expect(f.existsSync).not.toHaveBeenCalled()
})

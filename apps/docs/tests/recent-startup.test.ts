// @vitest-environment node
import { readFileSync } from 'node:fs'
import { basename } from 'node:path'
import { createContext, runInContext } from 'node:vm'
import ts from 'typescript'
import { expect, it, vi } from 'vitest'

const source = ts.createSourceFile(
  'docs-main.ts',
  readFileSync(new URL('../src/main/docs-main.ts', import.meta.url), 'utf8'),
  ts.ScriptTarget.Latest,
  true,
)

function fixture() {
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
    activeViewMenuState: () => ({}),
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

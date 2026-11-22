import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createContext, runInContext } from 'node:vm'
import ts from 'typescript'
import { afterEach, expect, it } from 'vitest'
import { getUiLang, isLang, normalizeLang, setUiLang, type Lang } from '@genoffice/i18n'
import { readAppSettings } from '../src/main/app-settings'

const shell = ts.createSourceFile(
  'index.ts',
  readFileSync(new URL('../src/main/index.ts', import.meta.url), 'utf8'),
  ts.ScriptTarget.Latest,
  true,
)
const docs = ts.createSourceFile(
  'docs-main.ts',
  readFileSync(new URL('../../docs/src/main/docs-main.ts', import.meta.url), 'utf8'),
  ts.ScriptTarget.Latest,
  true,
)
const currentLang = shell.statements.find(
  (node): node is ts.FunctionDeclaration =>
    ts.isFunctionDeclaration(node) && node.name?.text === 'currentLang',
)!
let ready: ts.Block | undefined
let languageHandler: ts.CallExpression | undefined
function visit(node: ts.Node): void {
  if (ts.isCallExpression(node) && node.expression.getText(shell) === 'app.whenReady().then') {
    const callback = node.arguments[0]
    if (callback && ts.isArrowFunction(callback) && ts.isBlock(callback.body)) ready = callback.body
  }
  ts.forEachChild(node, visit)
}
visit(shell)
function visitDocs(node: ts.Node): void {
  if (
    ts.isCallExpression(node) &&
    node.expression.getText(docs) === 'ipcMain.handle' &&
    node.arguments[0]?.getText(docs) === "'app:get-language'"
  )
    languageHandler = node
  ts.forEachChild(node, visitDocs)
}
visitDocs(docs)
const branch = ready!.statements.find(
  (node): node is ts.IfStatement =>
    ts.isIfStatement(node) && node.expression.getText(shell).includes('headlessArgv.kind'),
)
if (!branch || !languageHandler)
  throw new Error('Headless ready entry or Docs language bridge missing')
const headlessSteps = (branch.thenStatement as ts.Block).statements.map((node) =>
  node.getText(shell),
)
const directories: string[] = []
afterEach(() => {
  setUiLang('zh')
  for (const path of directories.splice(0)) rmSync(path, { recursive: true, force: true })
})

it.each([
  ['en', 'zh-CN', undefined, 'en'],
  ['fr', 'en-US', undefined, 'fr'],
  ['en', 'en-US', 'ja-JP', 'ja'],
  ['invalid', 'de-DE', undefined, 'de'],
  [undefined, 'en-GB', undefined, 'en'],
  ['zh', 'en-US', undefined, 'zh'],
] as const)(
  'initializes export language from saved=%s OS=%s override=%s',
  async (saved, locale, override, expected) => {
    const directory = mkdtempSync(join(tmpdir(), 'genoffice-headless-language-'))
    directories.push(directory)
    const settings = join(directory, 'app-settings.json')
    const original = JSON.stringify({ language: saved })
    writeFileSync(settings, original)
    setUiLang('zh')
    let getLanguage: (() => Lang) | undefined
    let exported: Lang | undefined
    const context = createContext({
      uiLang: null,
      process: { env: { GENOFFICE_LANG: override } },
      app: { getLocale: () => locale },
      APP_SETTINGS_PATH: () => settings,
      readAppSettings,
      isLang,
      normalizeLang,
      setUiLang,
      getUiLang,
      ipcMain: {
        handle: (_channel: string, handler: () => Lang) => {
          getLanguage = handler
        },
      },
      headlessArgv: { kind: 'ok' },
      runHeadlessExportEntry: async () => {
        exported = await getLanguage!()
      },
    })
    const execute = (code: string) =>
      runInContext(
        ts.transpileModule(code, {
          compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.CommonJS },
        }).outputText,
        context,
      )
    // only the headless branch runs: the startup steps before it change often
    // and have nothing to do with the language
    execute(currentLang.getText(shell))
    execute(languageHandler!.getText(docs))
    await execute(`(async () => { ${headlessSteps.join('\n')} })()`)
    expect(exported).toBe(expected)
    expect(readFileSync(settings, 'utf8')).toBe(original)
  },
)

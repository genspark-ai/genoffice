import { readFileSync } from 'node:fs'
import { runInNewContext } from 'node:vm'
import ts from 'typescript'
import { createI18n, type Lang } from '@genoffice/i18n'
import type { TabMenuStringKey } from '../../src/main/tab-menu'

// Read the production table without importing index.ts and starting Electron.
const source = ts.createSourceFile(
  'index.ts',
  readFileSync(new URL('../../src/main/index.ts', import.meta.url), 'utf8'),
  ts.ScriptTarget.Latest,
  true,
)
let table: ts.Expression | undefined
function findTable(node: ts.Node): void {
  if (
    ts.isVariableDeclaration(node) &&
    node.name.getText(source) === 'tMain' &&
    node.initializer &&
    ts.isCallExpression(node.initializer)
  )
    table = node.initializer.arguments[0]
  ts.forEachChild(node, findTable)
}
findTable(source)
if (!table || !ts.isObjectLiteralExpression(table))
  throw new Error('Shell main string table missing')
export const mainMenuStrings = runInNewContext(`(${table.getText(source)})`) as Record<
  Lang,
  Record<TabMenuStringKey, string>
>
export const translateMainMenu = createI18n(mainMenuStrings)
export const tm = (key: TabMenuStringKey): string => translateMainMenu('en', key)

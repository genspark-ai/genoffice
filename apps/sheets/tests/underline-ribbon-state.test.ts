/** @vitest-environment jsdom */
import { readFileSync } from 'node:fs'
import { URL as NodeURL } from 'node:url'
import {
  BooleanNumber,
  LogLevel,
  Univer,
  UniverInstanceType,
  type IStyleData,
} from '@univerjs/core'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import '@univerjs/sheets/facade'
import { FUniver } from '@univerjs/core/facade'
import { createElement, act } from 'react'
import * as jsxRuntime from 'react/jsx-runtime'
import { createRoot } from 'react-dom/client'
import ts from 'typescript'
import { describe, expect, it } from 'vitest'

import { selectionFormatEquals, toSelectionFormat } from '../src/renderer/selection-format'

// Render the production buttons, including their state expressions and handlers.
const source = readFileSync(new NodeURL('../src/renderer/ExcelShell.tsx', import.meta.url), 'utf8')
const tree = ts.createSourceFile(
  'ExcelShell.tsx',
  source,
  ts.ScriptTarget.Latest,
  true,
  ts.ScriptKind.TSX,
)
const buttons: string[] = []
function visit(node: ts.Node) {
  if (ts.isJsxElement(node) && node.openingElement.tagName.getText(tree) === 'button') {
    const text = node.getText(tree)
    if (/data-tip=\{t\('app(?:Double)?Underline'\)\}/.test(text)) buttons.push(text)
  }
  ts.forEachChild(node, visit)
}
visit(tree)
const compiled = ts.transpileModule(
  `export function Buttons({ selectionFormat, onCommand }) { const t = key => key; return <>${buttons.join('\n')}</> }`,
  { compilerOptions: { jsx: ts.JsxEmit.ReactJSX, module: ts.ModuleKind.CommonJS } },
).outputText
const exports: { Buttons?: (props: unknown) => unknown } = {}
new Function('require', 'exports', compiled)((id: string) => {
  if (id === 'react/jsx-runtime') return jsxRuntime
  throw new Error(id)
}, exports)

describe('underline ribbon echo', () => {
  it('preserves the first-cell echo for mixed ranges', () => {
    const univer = new Univer({ logLevel: LogLevel.SILENT })
    univer.registerPlugin(UniverSheetsPlugin)
    univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
      id: 'underline-echo',
      name: 'Underline',
      sheetOrder: ['s1'],
      sheets: {
        s1: {
          id: 's1',
          name: 'Sheet1',
          rowCount: 10,
          columnCount: 10,
          cellData: {
            0: {
              0: { v: 'single', s: { ul: { s: BooleanNumber.TRUE } } },
              1: { v: 'double', s: { ul: { s: BooleanNumber.TRUE, t: 10 } } },
              2: { v: 'off' },
              3: { v: 'single', s: { ul: { s: BooleanNumber.TRUE } } },
            },
          },
        },
      },
    })
    try {
      const sheet = FUniver.newAPI(univer).getActiveWorkbook()!.getActiveSheet()!
      for (const [range, underline, doubleUnderline] of [
        ['A1:D1', true, false],
        ['B1:D1', true, true],
        ['C1:D1', false, false],
      ] as const) {
        const format = toSelectionFormat(sheet.getRange(range).getCellStyleData() ?? {}, 'General')
        expect(format).toMatchObject({ underline, doubleUnderline })
      }
    } finally {
      univer.dispose()
    }
  })
  it('refreshes the selection when only the underline variant changes', () => {
    const single = toSelectionFormat({ ul: { s: BooleanNumber.TRUE } }, 'General')
    const double = toSelectionFormat({ ul: { s: BooleanNumber.TRUE, t: 10 } }, 'General')
    expect(selectionFormatEquals(single, double)).toBe(false)
    expect(selectionFormatEquals(double, single)).toBe(false)
    expect(selectionFormatEquals(double, { ...double })).toBe(true)
  })

  it('highlights exactly the matching button across selection changes and keeps command routing', () => {
    expect(buttons).toHaveLength(2)
    const container = document.createElement('div')
    const root = createRoot(container)
    const commands: string[] = []
    const cases: [IStyleData | null, boolean, boolean][] = [
      [{ ul: { s: BooleanNumber.TRUE } }, true, false],
      [{ ul: { s: BooleanNumber.TRUE, t: 10 } }, false, true],
      [{ ul: { s: BooleanNumber.FALSE, t: 10 } }, false, false],
      [{}, false, false],
      [null, false, false],
      [{ ul: { s: BooleanNumber.TRUE, t: 10 } }, false, true],
      [{ ul: { s: BooleanNumber.TRUE } }, true, false],
    ]
    try {
      for (const [style, single, double] of cases) {
        act(() =>
          root.render(
            createElement(exports.Buttons as never, {
              selectionFormat: style === null ? null : toSelectionFormat(style, 'General'),
              onCommand: (command: string) => commands.push(command),
            }),
          ),
        )
        const controls = [...container.querySelectorAll('button')]
        expect(controls.map((button) => button.classList.contains('is-active'))).toEqual([
          single,
          double,
        ])
      }
      act(() => container.querySelectorAll('button').forEach((button) => button.click()))
      expect(commands).toEqual(['underline', 'underline:double'])
    } finally {
      act(() => root.unmount())
    }
  })
})

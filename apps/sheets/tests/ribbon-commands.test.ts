import { describe, expect, it } from 'vitest'

import { EDITOR_ACTIVATED, HorizontalAlign } from '@univerjs/core'

import {
  handleRibbonCommand,
  parseStyleCommand,
  type RibbonCommandContext,
} from '../src/renderer/ribbon-actions'

describe('parseStyleCommand', () => {
  it('splits plain name:argument commands', () => {
    expect(parseStyleCommand('fill:#00aa55')).toEqual({
      name: 'fill',
      argument: '#00aa55',
      extra: '',
    })
    expect(parseStyleCommand('align:right')).toEqual({
      name: 'align',
      argument: 'right',
      extra: '',
    })
  })

  it('handles commands without an argument', () => {
    expect(parseStyleCommand('bold')).toEqual({ name: 'bold', argument: '', extra: '' })
  })

  it('keeps colons inside number-format patterns intact', () => {
    expect(parseStyleCommand('format:h:mm:ss AM/PM')).toEqual({
      name: 'format',
      argument: 'h:mm:ss AM/PM',
      extra: '',
    })
    expect(parseStyleCommand('format:hh:mm:ss')).toEqual({
      name: 'format',
      argument: 'hh:mm:ss',
      extra: '',
    })
  })

  it('still splits the extra segment for three-part commands', () => {
    expect(parseStyleCommand('border:all:#112233')).toEqual({
      name: 'border',
      argument: 'all',
      extra: '#112233',
    })
    expect(parseStyleCommand('cellprot:locked-on:hidden-off')).toEqual({
      name: 'cellprot',
      argument: 'locked-on',
      extra: 'hidden-off',
    })
    expect(parseStyleCommand('cellprot:locked-on:')).toEqual({
      name: 'cellprot',
      argument: 'locked-on',
      extra: '',
    })
    expect(parseStyleCommand('sort-custom:1:2a,3d')).toEqual({
      name: 'sort-custom',
      argument: '1',
      extra: '2a,3d',
    })
  })
})

/// Mirrors transformFacadeHorizontalAlignment in @univerjs/sheets' facade:
/// only these three values are accepted, and 'normal' means RIGHT — anything
/// else (such as a raw 'right') hits the throwing default branch.
function facadeHorizontalAlignment(value: string): HorizontalAlign {
  switch (value) {
    case 'left':
      return HorizontalAlign.LEFT
    case 'center':
      return HorizontalAlign.CENTER
    case 'normal':
      return HorizontalAlign.RIGHT
    default:
      throw new Error(`Invalid horizontal alignment: ${value}`)
  }
}

interface DispatchHarness {
  ctx: RibbonCommandContext
  model: { ht?: HorizontalAlign; numberFormat?: string }
  messages: string[]
}

function makeDispatchHarness(): DispatchHarness {
  const model: DispatchHarness['model'] = {}
  const messages: string[] = []
  const range = {
    getCellStyleData: () => ({}),
    setHorizontalAlignment: (value: string) => {
      model.ht = facadeHorizontalAlignment(value)
    },
    setNumberFormat: (pattern: string) => {
      model.numberFormat = pattern
    },
  }
  const workbook = {
    getActiveSheet: () => undefined,
    getActiveRange: () => range,
  }
  const ctx = {
    univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
    setMessage: (message: string) => {
      messages.push(message)
    },
  } as unknown as RibbonCommandContext
  return { ctx, model, messages }
}

describe('handleRibbonCommand alignment', () => {
  it.each([
    ['align:left', HorizontalAlign.LEFT],
    ['align:center', HorizontalAlign.CENTER],
    ['align:right', HorizontalAlign.RIGHT],
  ])('%s lands in the model', (command, expected) => {
    const { ctx, model, messages } = makeDispatchHarness()
    handleRibbonCommand(ctx, command)
    expect(model.ht).toBe(expected)
    // The dispatcher reports the applied-style message, not a caught error.
    expect(messages).toHaveLength(1)
    expect(messages[0]).not.toMatch(/Invalid horizontal alignment/)
  })
})

describe('handleRibbonCommand number format', () => {
  it('passes patterns containing colons through unclipped', () => {
    const { ctx, model } = makeDispatchHarness()
    handleRibbonCommand(ctx, 'format:h:mm:ss AM/PM')
    expect(model.numberFormat).toBe('h:mm:ss AM/PM')
  })
})

describe('use-in-formula cell guard', () => {
  function makeUseInFormulaHarness(cell: { formula?: string; value?: unknown }) {
    const messages: string[] = []
    const writes: unknown[] = []
    const cellRange = {
      getFormula: () => cell.formula ?? '',
      getValue: () => (cell.value === undefined ? null : cell.value),
      setValue: (value: unknown) => {
        writes.push(value)
      },
    }
    const active = { getRow: () => 0, getColumn: () => 0 }
    const worksheet = { getRange: () => cellRange }
    const workbook = { getActiveSheet: () => worksheet, getActiveRange: () => active }
    const runtime = { univerAPI: { getActiveWorkbook: () => workbook } }
    const setMessage = (message: string) => {
      messages.push(message)
    }
    const ctx = {
      univerRef: { current: runtime },
      setMessage,
      dataToolsContext: () => ({ univerRef: { current: runtime }, setMessage }),
    } as unknown as RibbonCommandContext
    return { ctx, writes, messages }
  }

  it('refuses to overwrite a cell holding a formula', () => {
    const { ctx, writes, messages } = makeUseInFormulaHarness({
      formula: '=SUM(A1:A3)*Rate',
      value: 42,
    })
    handleRibbonCommand(ctx, 'use-in-formula:Revenue')
    expect(writes).toHaveLength(0)
    expect(messages).toHaveLength(1)
  })

  it('refuses to overwrite a header label', () => {
    const { ctx, writes } = makeUseInFormulaHarness({ value: '销售额' })
    handleRibbonCommand(ctx, 'use-in-formula:销售额')
    expect(writes).toHaveLength(0)
  })

  it('writes the name into an empty cell', () => {
    const { ctx, writes } = makeUseInFormulaHarness({})
    handleRibbonCommand(ctx, 'use-in-formula:Revenue')
    expect(writes).toEqual([{ f: '=Revenue' }])
  })
})

describe('in-cell editor inline formatting', () => {
  function makeEditorHarness(editing: boolean) {
    const cellWrites: Record<string, unknown[]> = {
      fontColor: [],
      fontWeight: [],
      fontFamily: [],
      fontSize: [],
    }
    const inlineCommands: Array<{ id: string; params: unknown }> = []
    const range = {
      getCellStyleData: () => ({}),
      getFontSize: () => 11,
      setFontColor: (value: unknown) => cellWrites.fontColor!.push(value),
      setFontWeight: (value: unknown) => cellWrites.fontWeight!.push(value),
      setFontFamily: (value: unknown) => cellWrites.fontFamily!.push(value),
      setFontSize: (value: unknown) => cellWrites.fontSize!.push(value),
    }
    const workbook = {
      getActiveSheet: () => undefined,
      getActiveRange: () => range,
    }
    const contextService = {
      getContextValue: (key: string) => editing && key === EDITOR_ACTIVATED,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => workbook,
        executeCommand: (id: string, params?: unknown) => {
          inlineCommands.push({ id, params })
          return Promise.resolve(true)
        },
      },
      univer: { __getInjector: () => ({ get: () => contextService }) },
    }
    const ctx = {
      univerRef: { current: runtime },
      setMessage: () => {},
    } as unknown as RibbonCommandContext
    return { ctx, cellWrites, inlineCommands }
  }

  it('font color goes to the run-level inline command while editing', () => {
    const { ctx, cellWrites, inlineCommands } = makeEditorHarness(true)
    handleRibbonCommand(ctx, 'font-color:#C00000')
    expect(cellWrites.fontColor).toHaveLength(0)
    expect(inlineCommands).toEqual([
      { id: 'doc.command.set-inline-format-text-color', params: { value: '#C00000' } },
    ])
  })

  it('font color "auto" clears the run color while editing', () => {
    const { ctx, inlineCommands } = makeEditorHarness(true)
    handleRibbonCommand(ctx, 'font-color:auto')
    expect(inlineCommands).toEqual([
      { id: 'doc.command.set-inline-format-text-color', params: { value: null } },
    ])
  })

  it('font color still styles the whole cell when not editing', () => {
    const { ctx, cellWrites, inlineCommands } = makeEditorHarness(false)
    handleRibbonCommand(ctx, 'font-color:#C00000')
    expect(inlineCommands).toHaveLength(0)
    expect(cellWrites.fontColor).toEqual(['#C00000'])
  })

  it('bold toggle routes inline while editing, absolute :on stays cell-wide', () => {
    const { ctx, cellWrites, inlineCommands } = makeEditorHarness(true)
    handleRibbonCommand(ctx, 'bold')
    expect(inlineCommands).toEqual([
      { id: 'doc.command.set-inline-format-bold', params: undefined },
    ])
    expect(cellWrites.fontWeight).toHaveLength(0)
    handleRibbonCommand(ctx, 'bold:on')
    expect(cellWrites.fontWeight).toEqual(['bold'])
  })

  it('font family and size route inline with their values while editing', () => {
    const { ctx, cellWrites, inlineCommands } = makeEditorHarness(true)
    handleRibbonCommand(ctx, 'font-family:Arial')
    handleRibbonCommand(ctx, 'font-size:14')
    expect(inlineCommands).toEqual([
      { id: 'doc.command.set-inline-format-font-family', params: { value: 'Arial' } },
      { id: 'doc.command.set-inline-format-fontsize', params: { value: 14 } },
    ])
    expect(cellWrites.fontFamily).toHaveLength(0)
    expect(cellWrites.fontSize).toHaveLength(0)
  })
})

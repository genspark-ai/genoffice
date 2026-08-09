import { describe, expect, it, vi } from 'vitest'

import { HorizontalAlign } from '@univerjs/core'

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

describe('handleRibbonCommand clear contents', () => {
  it("delegates range discovery to Univer's authoritative selection service", () => {
    const calls: { id: string; params: unknown }[] = []
    const firstRange = {
      getRange: () => ({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }),
    }
    const secondRange = {
      getRange: () => ({ startRow: 2, endRow: 4, startColumn: 3, endColumn: 3 }),
    }
    const ctx = {
      univerRef: {
        current: {
          univerAPI: {
            getActiveWorkbook: () => ({
              getActiveSheet: () => ({
                getSelection: () => ({
                  getActiveRange: () => firstRange,
                  getActiveRangeList: () => [firstRange, secondRange],
                }),
              }),
            }),
            executeCommand: (id: string, params: unknown) => {
              calls.push({ id, params })
              return Promise.resolve(true)
            },
          },
        },
      },
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'clear-contents')

    expect(calls).toEqual([{ id: 'sheet.command.clear-selection-content', params: undefined }])
  })

  it('does not depend on the facade active-range list shape', () => {
    const calls: { id: string; params: unknown }[] = []
    const primary = { getRange: () => ({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 4 }) }
    const ctx = {
      univerRef: {
        current: {
          univerAPI: {
            getActiveWorkbook: () => ({
              getActiveSheet: () => ({
                getSelection: () => ({
                  getActiveRange: () => primary,
                  getActiveRangeList: () => [],
                }),
              }),
            }),
            executeCommand: (id: string, params: unknown) => {
              calls.push({ id, params })
              return Promise.resolve(true)
            },
          },
        },
      },
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'clear-contents')

    expect(calls).toEqual([{ id: 'sheet.command.clear-selection-content', params: undefined }])
  })

  it('lets Univer reject a clear when no selection exists', () => {
    const executeCommand = vi.fn<(id: string, params: unknown) => Promise<boolean>>(() =>
      Promise.resolve(true),
    )
    const ctx = {
      univerRef: {
        current: {
          univerAPI: {
            getActiveWorkbook: () => ({ getActiveSheet: () => ({ getSelection: () => null }) }),
            executeCommand,
          },
        },
      },
    } as unknown as RibbonCommandContext

    handleRibbonCommand(ctx, 'clear-contents')

    expect(executeCommand).toHaveBeenCalledWith('sheet.command.clear-selection-content')
  })
})

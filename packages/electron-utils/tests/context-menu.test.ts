import { describe, expect, it, vi } from 'vitest'

const electronMock = vi.hoisted(() => ({ popup: vi.fn() }))

vi.mock('electron', () => ({
  Menu: {
    buildFromTemplate: vi.fn(() => ({ popup: electronMock.popup })),
  },
}))

import {
  appMenuLabels,
  buildContextMenuItems,
  contextMenuLabels,
  installContextMenu,
  setContextMenuInterceptor,
} from '../src/index'

const { popup } = electronMock

const labels = contextMenuLabels('en')

const flags = (
  overrides: Partial<{ canCut: boolean; canCopy: boolean; canPaste: boolean }> = {},
) => ({
  canUndo: false,
  canRedo: false,
  canCut: true,
  canCopy: true,
  canPaste: true,
  canDelete: false,
  canSelectAll: true,
  canEditRichly: false,
  ...overrides,
})

const base = {
  isEditable: false,
  selectionText: '',
  misspelledWord: '',
  dictionarySuggestions: [] as string[],
  editFlags: flags(),
}

describe('buildContextMenuItems', () => {
  it('builds the full edit set for editable targets', () => {
    const items = buildContextMenuItems({ ...base, isEditable: true }, labels)
    expect(items).toEqual([
      { action: 'cut', label: 'Cut', enabled: true },
      { action: 'copy', label: 'Copy', enabled: true },
      { action: 'paste', label: 'Paste', enabled: true },
      { type: 'separator' },
      { action: 'selectAll', label: 'Select All', enabled: true },
    ])
  })

  it('disables items per editFlags', () => {
    const items = buildContextMenuItems(
      { ...base, isEditable: true, editFlags: flags({ canCut: false, canPaste: false }) },
      labels,
    )
    expect(items).toContainEqual({ action: 'cut', label: 'Cut', enabled: false })
    expect(items).toContainEqual({ action: 'paste', label: 'Paste', enabled: false })
  })

  it('prepends spelling suggestions when a word is misspelled', () => {
    const items = buildContextMenuItems(
      {
        ...base,
        isEditable: true,
        misspelledWord: 'helo',
        dictionarySuggestions: ['hello', 'halo'],
      },
      labels,
    )
    expect(items.slice(0, 3)).toEqual([
      { action: 'replaceMisspelling', label: 'hello' },
      { action: 'replaceMisspelling', label: 'halo' },
      { type: 'separator' },
    ])
  })

  it('skips the suggestion block when there are no suggestions', () => {
    const items = buildContextMenuItems(
      { ...base, isEditable: true, misspelledWord: 'zqx' },
      labels,
    )
    expect(items[0]).toEqual({ action: 'cut', label: 'Cut', enabled: true })
  })

  it('offers copy/select-all for non-editable selections', () => {
    const items = buildContextMenuItems({ ...base, selectionText: 'some text' }, labels)
    expect(items).toEqual([
      { action: 'copy', label: 'Copy', enabled: true },
      { action: 'selectAll', label: 'Select All', enabled: true },
    ])
  })

  it('returns nothing for non-editable targets without selection', () => {
    expect(buildContextMenuItems(base, labels)).toEqual([])
    expect(buildContextMenuItems({ ...base, selectionText: '   ' }, labels)).toEqual([])
  })
})

describe('contextMenuLabels', () => {
  it('covers every language and falls back to English', () => {
    expect(contextMenuLabels('zh').copy).toBe('复制')
    expect(contextMenuLabels('zh-TW').paste).toBe('貼上')
    expect(contextMenuLabels('xx')).toEqual(contextMenuLabels('en'))
  })

  it('falls back to the base language for regional variants', () => {
    expect(contextMenuLabels('pt-BR').copy).toBe(contextMenuLabels('pt').copy)
    expect(contextMenuLabels('zh-Hant').copy).toBe(contextMenuLabels('zh').copy)
    expect(contextMenuLabels('de-AT').copy).toBe(contextMenuLabels('de').copy)
    expect(appMenuLabels('pt-BR').window).toBe(appMenuLabels('pt').window)
    expect(appMenuLabels('zh-Hant').window).toBe(appMenuLabels('zh').window)
  })
})

describe('image targets', () => {
  const image = { ...base, mediaType: 'image' as const, srcURL: 'data:image/png;base64,AAAA' }

  it('offers view / copy / save on a non-editable image', () => {
    expect(buildContextMenuItems(image, labels)).toEqual([
      { action: 'viewImage', label: 'View Image' },
      { action: 'copyImage', label: 'Copy Image' },
      { action: 'saveImageAs', label: 'Save Image As…' },
    ])
  })

  it('puts the image block ahead of the edit set inside an editor', () => {
    const items = buildContextMenuItems({ ...image, isEditable: true }, labels)
    expect(items.slice(0, 4)).toEqual([
      { action: 'viewImage', label: 'View Image' },
      { action: 'copyImage', label: 'Copy Image' },
      { action: 'saveImageAs', label: 'Save Image As…' },
      { type: 'separator' },
    ])
    expect(items).toHaveLength(9)
  })

  it('ignores images without a source', () => {
    expect(buildContextMenuItems({ ...image, srcURL: '' }, labels)).toEqual([])
  })
})

describe('installContextMenu interceptor failures', () => {
  it('falls back to the native menu when the interceptor rejects', async () => {
    // Electron ignores the promise an async listener returns, so a rejection here
    // is an unhandled rejection in the main process
    const appOn = vi.fn()
    const app = { on: appOn } as unknown as Parameters<typeof installContextMenu>[0]
    const menuListeners: Array<(event: unknown, params: unknown) => void> = []
    const contents = {
      id: 7,
      on: vi.fn((_event: string, listener: (event: unknown, params: unknown) => void) =>
        menuListeners.push(listener),
      ),
    }
    const failing = vi.fn(async () => {
      throw new Error('renderer is gone')
    })
    setContextMenuInterceptor(app, contents as never, failing)
    installContextMenu(app, () => labels)
    const created = appOn.mock.calls[0]![1] as (event: unknown, contents: unknown) => void
    created({}, contents)
    const rejected: unknown[] = []
    const onRejection = (err: unknown): void => void rejected.push(err)
    process.on('unhandledRejection', onRejection)
    try {
      menuListeners[0]!(null, { ...base, selectionText: 'selected' })
      await new Promise((r) => setTimeout(r, 10))
    } finally {
      process.off('unhandledRejection', onRejection)
    }
    expect(rejected).toEqual([])
    // the native menu still pops, so a right-click is never dead
    expect(popup).toHaveBeenCalledTimes(1)
  })
})

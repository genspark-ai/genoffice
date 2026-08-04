import { describe, expect, it } from 'vitest'

import { buildContextMenuItems, contextMenuLabels } from '../src/index'

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
})

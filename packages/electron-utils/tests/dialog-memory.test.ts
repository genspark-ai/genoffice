import { join } from 'node:path'

import { describe, expect, it, vi } from 'vitest'

import { saveAsSuggestion, showOpenDialogWithMemory, showSaveDialogWithMemory } from '../src/index'

import type { Dialog } from 'electron'

function fakeDialog(overrides: Partial<Dialog> = {}): Dialog {
  return {
    showOpenDialog: vi.fn().mockResolvedValue({ canceled: true, filePaths: [] }),
    showSaveDialog: vi.fn().mockResolvedValue({ canceled: true, filePath: '' }),
    ...overrides,
  } as unknown as Dialog
}

const pickedOpen = (paths: string[]) =>
  vi.fn().mockResolvedValue({ canceled: false, filePaths: paths })
const pickedSave = (path: string) => vi.fn().mockResolvedValue({ canceled: false, filePath: path })

describe('showOpenDialogWithMemory', () => {
  it('passes options through unchanged before any pick was made', async () => {
    const dialog = fakeDialog()
    await showOpenDialogWithMemory(dialog, undefined, { properties: ['openFile'] })
    expect(dialog.showOpenDialog).toHaveBeenCalledWith({ properties: ['openFile'] })
  })

  it('remembers the picked directory and injects it as defaultPath next time', async () => {
    const dialog = fakeDialog({ showOpenDialog: pickedOpen([join('/work', 'report.docx')]) })
    await showOpenDialogWithMemory(dialog, undefined, {})
    await showOpenDialogWithMemory(dialog, undefined, {})
    expect(dialog.showOpenDialog).toHaveBeenLastCalledWith({ defaultPath: '/work' })
  })

  it('forwards the parent window when given', async () => {
    const dialog = fakeDialog()
    const parent = { id: 1 } as never
    await showOpenDialogWithMemory(dialog, parent, {})
    expect(dialog.showOpenDialog).toHaveBeenCalledWith(parent, {})
  })

  it('remembers the selected directory itself for openDirectory pickers', async () => {
    const dialog = fakeDialog({ showOpenDialog: pickedOpen(['/exports/images']) })
    await showOpenDialogWithMemory(dialog, undefined, { properties: ['openDirectory'] })
    await showOpenDialogWithMemory(dialog, undefined, {})
    expect(dialog.showOpenDialog).toHaveBeenLastCalledWith({ defaultPath: '/exports/images' })
  })

  it('keeps an explicit absolute defaultPath untouched', async () => {
    const dialog = fakeDialog({ showOpenDialog: pickedOpen([join('/work', 'a.docx')]) })
    await showOpenDialogWithMemory(dialog, undefined, {})
    await showOpenDialogWithMemory(dialog, undefined, { defaultPath: '/elsewhere/b.docx' })
    expect(dialog.showOpenDialog).toHaveBeenLastCalledWith({ defaultPath: '/elsewhere/b.docx' })
  })

  it('does not remember canceled picks', async () => {
    const dialog = fakeDialog()
    await showOpenDialogWithMemory(dialog, undefined, {})
    await showOpenDialogWithMemory(dialog, undefined, {})
    expect(dialog.showOpenDialog).toHaveBeenLastCalledWith({})
  })
})

describe('showSaveDialogWithMemory', () => {
  it('anchors a bare file-name suggestion in the remembered directory', async () => {
    const dialog = fakeDialog({ showSaveDialog: pickedSave(join('/work', 'deck.pptx')) })
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'deck.pptx' })
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'deck2.pptx' })
    expect(dialog.showSaveDialog).toHaveBeenLastCalledWith({
      defaultPath: join('/work', 'deck2.pptx'),
    })
  })

  it('shares the remembered directory between save and open dialogs', async () => {
    const dialog = fakeDialog({ showSaveDialog: pickedSave(join('/work', 'deck.pptx')) })
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'deck.pptx' })
    await showOpenDialogWithMemory(dialog, undefined, {})
    expect(dialog.showOpenDialog).toHaveBeenCalledWith({ defaultPath: '/work' })
  })

  it('keeps an explicit absolute defaultPath untouched', async () => {
    const dialog = fakeDialog({ showSaveDialog: pickedSave(join('/work', 'a.pdf')) })
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: '/docs/tab.pdf' })
    expect(dialog.showSaveDialog).toHaveBeenCalledWith({ defaultPath: '/docs/tab.pdf' })
  })

  it('isolates memory between different dialog instances', async () => {
    const first = fakeDialog({ showOpenDialog: pickedOpen([join('/work', 'a.docx')]) })
    const second = fakeDialog()
    await showOpenDialogWithMemory(first, undefined, {})
    await showOpenDialogWithMemory(second, undefined, {})
    expect(second.showOpenDialog).toHaveBeenCalledWith({})
  })

  it('anchors a bare file-name suggestion in the fallback dir when nothing is remembered', async () => {
    const dialog = fakeDialog()
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'deck.pptx' }, '/default/save')
    expect(dialog.showSaveDialog).toHaveBeenCalledWith({
      defaultPath: join('/default/save', 'deck.pptx'),
    })
  })

  it('prefers the remembered directory over the fallback dir', async () => {
    const dialog = fakeDialog({ showSaveDialog: pickedSave(join('/work', 'deck.pptx')) })
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'deck.pptx' })
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'b.pptx' }, '/default/save')
    expect(dialog.showSaveDialog).toHaveBeenLastCalledWith({
      defaultPath: join('/work', 'b.pptx'),
    })
  })

  it('keeps an explicit absolute defaultPath untouched even with a fallback dir', async () => {
    const dialog = fakeDialog()
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: '/docs/tab.pdf' }, '/default')
    expect(dialog.showSaveDialog).toHaveBeenCalledWith({ defaultPath: '/docs/tab.pdf' })
  })
})

describe('saveAsSuggestion', () => {
  it('suggests the source document folder with the new name (Word parity)', () => {
    expect(saveAsSuggestion(join('/work', 'report.docx'), 'report.docx')).toBe(
      join('/work', 'report.docx'),
    )
    expect(saveAsSuggestion(join('/work', 'report.docx'), 'copy.docx')).toBe(
      join('/work', 'copy.docx'),
    )
  })

  it('falls back to the bare name for a document that was never on disk', () => {
    expect(saveAsSuggestion(null, 'Untitled.docx')).toBe('Untitled.docx')
    expect(saveAsSuggestion(undefined, 'Untitled.docx')).toBe('Untitled.docx')
    expect(saveAsSuggestion('', 'Untitled.pptx')).toBe('Untitled.pptx')
  })

  it('beats the remembered directory when threaded through the save dialog', async () => {
    const dialog = fakeDialog({ showSaveDialog: pickedSave(join('/elsewhere', 'x.docx')) })
    // a pick in /elsewhere seeds the remembered directory…
    await showSaveDialogWithMemory(dialog, undefined, { defaultPath: 'x.docx' })
    // …but Save As of a document living in /work still opens in /work
    await showSaveDialogWithMemory(dialog, undefined, {
      defaultPath: saveAsSuggestion(join('/work', 'report.docx'), 'report.docx'),
    })
    expect(dialog.showSaveDialog).toHaveBeenLastCalledWith({
      defaultPath: join('/work', 'report.docx'),
    })
  })
})

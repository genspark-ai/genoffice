import { readFileSync } from 'node:fs'

import { describe, expect, it } from 'vitest'

import { IPC_CHANNELS } from '../src/shared/ipc-channels'

/**
 * The delimited import (a .csv/.tsv opened as a workbook) reports its
 * progress and can be cancelled. The pieces live on both sides of the IPC
 * boundary, and a break in the middle is invisible in any one of them: the
 * main process can report progress nobody listens for, the renderer can
 * listen on a channel the main process never sends on, and the cancel can
 * reach a registry that the import never registered with.
 *
 * These assertions are on the source, not on behaviour — `sheets-main.ts`
 * touches `electron.app` at import time, so it cannot be imported here — but
 * they are the seams where the wiring goes missing.
 */
const mainSource = readFileSync(new URL('../src/main/sheets-main.ts', import.meta.url), 'utf8')
const preloadSource = readFileSync(new URL('../src/preload/index.ts', import.meta.url), 'utf8')
const appSource = readFileSync(new URL('../src/renderer/App.tsx', import.meta.url), 'utf8')
const hookSource = readFileSync(
  new URL('../src/renderer/import-progress.ts', import.meta.url),
  'utf8',
)

describe('the delimited import progress wiring', () => {
  it('sends the progress on the channel the preload listens on', () => {
    expect(mainSource).toContain(`send(IPC_CHANNELS.importProgress`)
    expect(preloadSource).toContain(`ipcRenderer.on(IPC_CHANNELS.importProgress`)
    // the channel name itself is the contract between the two sides
    expect(IPC_CHANNELS.importProgress).toBe('workbook:import-progress')
    expect(IPC_CHANNELS.cancelImport).toBe('workbook:cancel-import')
  })

  it('registers the import with the registry the cancel reaches', () => {
    expect(mainSource).toContain('pendingImports.start(owner)')
    expect(mainSource).toContain('pendingImports.cancel(event.sender.id)')
    // the controller the import runs under is the one the registry hands out
    expect(mainSource).toContain('signal: importController.signal')
  })

  it('hands the progress and the signal to the conversion', () => {
    expect(mainSource).toContain('options?.import')
    expect(mainSource).toContain('onProgress: reportProgress')
  })

  it('subscribes the renderer and gives it a way out', () => {
    expect(appSource).toContain('useImportProgress')
    expect(appSource).toContain('cancelImport')
    // the renderer unsubscribes, so a re-render does not stack listeners
    expect(hookSource).toContain('onImportProgress')
    expect(hookSource).toContain('stop?.()')
  })

  it('says the import is over when it ends', () => {
    // Without this the bar outlives its import: the next workbook would open
    // showing the previous file's percentage, with a cancel button that has
    // nothing to cancel.
    expect(mainSource).toContain('reportProgress(null)')
    expect(preloadSource).toContain('listener(null)')
  })

  it('opens nothing when the user cancelled the import', () => {
    // A cancelled import is the user's own doing, not a failed open: it must
    // not surface as an error the user has to read.
    expect(mainSource).toContain('pendingImports.wasCancelled(importController)')
    expect(appSource).toContain('appImportCanceled')
  })
})

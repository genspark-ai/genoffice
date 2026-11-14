/**
 * The print subsystem lifted out of App.tsx (#2461's split of the giant
 * renderer components): the dialog's fields, the driver option groups one
 * queue advertises, and the print path itself. App wires the cross-boundary
 * pieces in; PrintDialog.tsx renders the returned object verbatim, so the
 * names here are the names the dialog reads.
 */
import { useRef, useState } from 'react'
// legacy build: same reasoning as App.tsx — embedded font parsing needs it
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { filterPageParity, maySpoolOriginal, printPdf, type PageParity } from './print'
import { readFileBytes } from './file-bytes'
import { DOC_OPTS, parsePageRanges } from './view-config'
import type { PrintDuplex, PrinterInfo, PrinterOption } from '../shared/ipc'

/** the escape hatch: hand the job to the OS print dialog instead of printing here */
export const PRINT_SYSTEM_DIALOG = '\u0000system-dialog'
export const PRINT_DUPLEX_MODES: PrintDuplex[] = ['simplex', 'longEdge', 'shortEdge']

const IS_MAC = navigator.platform.toLowerCase().includes('mac')

export interface PrintJobInput {
  /** The flushed file on disk; the print re-reads it so the destroyed old doc never matters */
  filePath: string
  pageCount: number
  currentPage: number
  /** Saves pending edits first; undefined = the save failed and the job is off */
  flushThen: (job: () => Promise<void>) => Promise<void | undefined>
  opFailed: (error: string) => void
  /** The on-disk file's password; CUPS cannot open encrypted files, so those raster */
  diskPasswordRef: { current: string | null }
}

export function usePrintJob(input: PrintJobInput) {
  const { filePath, pageCount, currentPage, flushThen, opFailed, diskPasswordRef } = input
  const [printDlg, setPrintDlg] = useState(false)
  const [printMode, setPrintMode] = useState<'all' | 'current' | 'custom'>('all')
  const [printInput, setPrintInput] = useState('')
  const [printInvalid, setPrintInvalid] = useState(false)
  const [printParity, setPrintParity] = useState<PageParity>('all')
  const [printers, setPrinters] = useState<PrinterInfo[]>([])
  const [printDevice, setPrintDevice] = useState(IS_MAC ? PRINT_SYSTEM_DIALOG : '')
  const [printCopies, setPrintCopies] = useState(1)
  const [printCollate, setPrintCollate] = useState(true)
  const [printDuplex, setPrintDuplex] = useState<PrintDuplex>('simplex')
  // the chosen printer's driver option groups (quality, color, media type) and
  // what the user picked from them; empty for the system dialog and on
  // platforms where the CUPS route does not run
  const [driverOptions, setDriverOptions] = useState<PrinterOption[]>([])
  // bumped on every printer switch: a slow printerOptions() answer from an
  // earlier queue must not overwrite the current queue's groups
  const driverOptionsRequest = useRef(0)
  const [driverChoices, setDriverChoices] = useState<Record<string, string>>({})
  const [printing, setPrinting] = useState(false)
  const printCurrentPage = Math.max(1, Math.min(currentPage, pageCount))

  const openPrintDlg = () => {
    setPrintMode('all')
    setPrintInput(String(printCurrentPage))
    setPrintInvalid(false)
    setPrintParity('all')
    setPrintDlg(true)
    window.pdfApi
      .listPrinters()
      .then(setPrinters)
      .catch(() => undefined)
  }

  /** Print: save first (markups/forms/page ops all into the file), then reload from the file to render, avoiding a destroyed old doc */
  // Synchronous re-entry guard: the menu accelerator and the renderer's own ⌘P can both
  // fire, and the `printing` state only updates after the async flush has started
  const printBusyRef = useRef(false)
  // `pages` are 1-based file pages; the flush compacts the file to the visible
  // order first, so visible page v prints as file page v.
  const printDoc = async (pages?: number[]) => {
    if (printBusyRef.current) return
    printBusyRef.current = true
    const job =
      printDevice === PRINT_SYSTEM_DIALOG
        ? undefined
        : {
            deviceName: printDevice,
            copies: printCopies,
            collate: printCollate,
            duplex: printDuplex,
            driverOptions: driverOptions
              .filter(
                (g) =>
                  driverChoices[g.name] !== undefined && driverChoices[g.name] !== g.defaultValue,
              )
              .map((g) => ({ name: g.name, value: driverChoices[g.name]! })),
          }
    try {
      await flushThen(async () => {
        setPrinting(true)
        try {
          if (job && IS_MAC && maySpoolOriginal(diskPasswordRef.current)) {
            // macOS named printer: lp spools the flushed file itself — the
            // vector original reaches the driver untouched, instead of a
            // 150-200 DPI raster pass (genoffice#2347's route, quality half).
            // An encrypted file stays encrypted on disk, which CUPS' pdftopdf
            // cannot open, so those fall through to the raster path below.
            const r = await window.pdfApi.printFile({ ...job, pages, path: filePath })
            if (!r.ok && r.error) opFailed(r.error)
            return
          }
          const data = await readFileBytes(window.pdfApi, filePath)
          const pdoc = await getDocument({ data, ...DOC_OPTS }).promise
          try {
            await printPdf(pdoc, pages, job)
          } finally {
            void pdoc.loadingTask.destroy()
          }
        } catch (err) {
          opFailed(err instanceof Error ? err.message : String(err))
        } finally {
          setPrinting(false)
        }
      })
    } finally {
      printBusyRef.current = false
    }
  }

  /** Print dialog confirm: visible page numbers → 1-based file pages */
  const confirmPrint = () => {
    const pages =
      printMode === 'all'
        ? Array.from({ length: pageCount }, (_x, i) => i + 1)
        : printMode === 'current'
          ? [printCurrentPage]
          : parsePageRanges(printInput, pageCount)
    const subset = pages ? filterPageParity(pages, printParity) : []
    if (subset.length === 0) {
      setPrintInvalid(true)
      return
    }
    setPrintDlg(false)
    void printDoc(printMode === 'all' && printParity === 'all' ? undefined : subset)
  }

  return {
    printing,
    printDlg,
    setPrintDlg,
    printMode,
    setPrintMode,
    printInput,
    setPrintInput,
    printInvalid,
    setPrintInvalid,
    printParity,
    setPrintParity,
    printers,
    printDevice,
    setPrintDevice,
    printCopies,
    setPrintCopies,
    printCollate,
    setPrintCollate,
    printDuplex,
    setPrintDuplex,
    driverOptions,
    setDriverOptions,
    driverOptionsRequest,
    driverChoices,
    setDriverChoices,
    printCurrentPage,
    openPrintDlg,
    confirmPrint,
  }
}

export type PrintJobApi = ReturnType<typeof usePrintJob>

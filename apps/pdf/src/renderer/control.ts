/** Request/reply the shell relays from `genoffice open --page` and `genoffice selection`. */
export type ControlRequest =
  { cmd: 'goto'; target: { kind: string; page?: number } } | { cmd: 'selection' }

export type ControlReply =
  | { status: 'ok'; result: Record<string, unknown> }
  | { status: 'not_ready' }
  | {
      status: 'error'
      error: { reason: string; message: string; detail?: Record<string, unknown> }
    }

export interface PdfControlState {
  loaded: boolean
  pageCount: number
  currentPage: number
  scrollToPage: (page: number) => void
  selectedText: () => string
}

export function handlePdfControl(req: ControlRequest, state: PdfControlState): ControlReply {
  if (!state.loaded || state.pageCount === 0) return { status: 'not_ready' }
  if (req.cmd === 'selection') {
    const text = state.selectedText()
    return { status: 'ok', result: { page: state.currentPage, ...(text ? { text } : {}) } }
  }
  const { page } = req.target
  if (page === undefined || !Number.isInteger(page) || page < 1 || page > state.pageCount) {
    return {
      status: 'error',
      error: {
        reason: 'out_of_range',
        message: `page ${page} is out of range (the document has ${state.pageCount} pages)`,
        detail: { valid_range: `1-${state.pageCount}` },
      },
    }
  }
  state.scrollToPage(page)
  return { status: 'ok', result: { page } }
}

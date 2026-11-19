/// <reference types="vite/client" />

import type { PdfApi } from '../shared/ipc'

declare global {
  interface Window {
    pdfApi: PdfApi
  }
}

export {}

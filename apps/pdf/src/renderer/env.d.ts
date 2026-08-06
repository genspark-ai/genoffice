/// <reference types="vite/client" />

import type { PdfApi } from '../shared/ipc'
import type { ProjectApi } from '@genoffice/project-store'

declare global {
  interface Window {
    pdfApi: PdfApi
    projectApi: ProjectApi
  }
}

export {}

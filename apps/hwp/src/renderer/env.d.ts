/// <reference types="vite/client" />

import type { HwpApi } from '../shared/ipc'

declare global {
  interface Window {
    hwpApi: HwpApi
  }
}

export {}

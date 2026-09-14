declare module '*.md?raw' {
  const content: string
  export default content
}

import type { DesktopApi } from '../shared/desktop-api'
import type { ProjectApi } from '@genoffice/project-store'

declare global {
  interface Window {
    readonly desktopApi: DesktopApi
    readonly projectApi: ProjectApi
  }
}

export {}

/// <reference types="vite/client" />

import type { DesktopApi } from '../shared/desktop-api'

declare global {
  interface Window {
    desktop: DesktopApi
  }
}

export {}

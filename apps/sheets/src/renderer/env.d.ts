declare module '*.md?raw' {
  const content: string
  export default content
}

import type { DesktopApi } from '../shared/desktop-api'
import type { ProjectApi } from '@genoffice/project-store'

import type { FilesPaneApi } from '@genoffice/ui'

declare global {
  interface Window {
    readonly desktopApi: DesktopApi
    readonly filesPaneApi: FilesPaneApi
    readonly projectApi: ProjectApi
  }
}

export {}

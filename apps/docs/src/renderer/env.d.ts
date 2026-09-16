/// <reference types="vite/client" />

import type { DesktopApi } from '../shared/ipc'
import type { ProjectApi } from '@genoffice/project-store'

import type { FilesPaneApi } from '@genoffice/ui'

declare global {
  interface Window {
    desktop: DesktopApi
    filesPaneApi: FilesPaneApi
    projectApi: ProjectApi
  }
}

export {}

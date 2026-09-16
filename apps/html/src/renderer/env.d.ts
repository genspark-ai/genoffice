/// <reference types="vite/client" />

import type { ProjectApi } from '@genoffice/project-store'
import type { HtmlApi } from '../shared/ipc'

import type { FilesPaneApi } from '@genoffice/ui'

declare global {
  interface Window {
    htmlApi: HtmlApi
    filesPaneApi: FilesPaneApi
    projectApi?: Pick<ProjectApi, 'resolveChat' | 'appendChat' | 'loadChat' | 'rebindChat'>
  }
}

export {}

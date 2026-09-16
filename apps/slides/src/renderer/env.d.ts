/// <reference types="vite/client" />
import type { SlidesApi } from '../shared/ipc'
import type { ProjectApi } from '@genoffice/project-store'

import type { FilesPaneApi } from '@genoffice/ui'

declare global {
  interface Window {
    slidesApi: SlidesApi
    filesPaneApi: FilesPaneApi
    projectApi: ProjectApi
  }
}

export {}

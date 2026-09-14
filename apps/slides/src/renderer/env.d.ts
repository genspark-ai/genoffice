/// <reference types="vite/client" />

import type { SlidesApi } from '../shared/ipc'
import type { ProjectApi } from '@genoffice/project-store'

declare global {
  interface Window {
    desktop: SlidesApi
    slidesApi: SlidesApi
    projectApi: ProjectApi
  }
}

export {}

/// <reference types="vite/client" />
import type { SlidesApi, DesktopFilesApi } from '../shared/ipc'
import type { ProjectApi } from '@genoffice/project-store'

declare global {
  interface Window {
    slidesApi: SlidesApi
    /** slides renderer file operations (the shell mounts a superset) */
    desktop: DesktopFilesApi
    projectApi: ProjectApi
  }
}

export {}

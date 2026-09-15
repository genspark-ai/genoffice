import type { GenOfficeEmbedEvent } from './embed-bridge'

declare global {
  interface Window {
    dataflareOfficeBridge?: {
      postEvent(event: GenOfficeEmbedEvent): void
      isEmbedded: boolean
      getRevision(): string
    }
  }
}

export {}

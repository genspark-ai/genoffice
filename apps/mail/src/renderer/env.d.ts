/// <reference types="vite/client" />

import type { ComposeApi } from '../shared/compose-channels'
import type { MailApi } from '../shared/ipc'
import type { QuarantineApi } from '../shared/quarantine-channels'

declare global {
  interface Window {
    mailApi: MailApi
    /** Compose-session bridge (Lane C); absent in unit tests and old preloads. */
    mailCompose?: ComposeApi
    /** Quarantine bridge (H5); absent in unit tests and old preloads. */
    mailQuarantine?: QuarantineApi
  }
}

export {}

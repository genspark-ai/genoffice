import { openOle2Container, type Ole2Limits } from '@genoffice/doc2docx/compound'

import { LegacyPptError } from './errors'
import { ENCRYPTED_CURRENT_USER_TOKEN, currentUserInfo } from './ppt'

/** Caps mirror the other legacy importers: fail fast, never walk gigabytes. */
const PPT_LIMITS: Ole2Limits = {
  maxBytes: 256 * 1024 * 1024,
  maxStreamBytes: 128 * 1024 * 1024,
}

/**
 * Open a legacy .ppt container. The OLE2/CFB parsing itself is the shared
 * opener in @genoffice/doc2docx (one copy across the legacy importers, with
 * the 8-byte signature and 512-byte header checks); this layer adds the
 * PowerPoint-specific checks: the presentation stream exists and the deck is
 * not password-protected (CurrentUserAtom headerToken 0xF3D1C4DF).
 */
export function openPptContainer(bytes: Uint8Array): {
  document: Uint8Array
  currentUser: Uint8Array | null
} {
  const opened = openOle2Container(bytes, PPT_LIMITS)
  if (!opened.ok) throw new LegacyPptError(opened.code, opened.detail)
  const container = opened.container
  const document = container.stream('PowerPoint Document')
  if (!document) {
    throw new LegacyPptError(
      'missing_stream',
      'the OLE2 container has no PowerPoint Document stream',
    )
  }
  const currentUser = container.stream('Current User') ?? null
  if (currentUser) {
    const info = currentUserInfo(currentUser)
    // an encrypted deck still has a readable Current User stream; the
    // presentation stream is encrypted, so fail before walking it
    if (info?.headerToken === ENCRYPTED_CURRENT_USER_TOKEN) {
      throw new LegacyPptError('encrypted')
    }
  }
  return { document, currentUser }
}

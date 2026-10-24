/**
 * Password-protected (ECMA-376 encrypted) docx support. The single integration
 * point for the third-party crypto implementation (officecrypto-tool): if it
 * ever needs replacing, only this module changes.
 *
 * An encrypted docx is not a zip: Word/WPS repackage it as a CFB (OLE2)
 * container holding an EncryptionInfo stream (parameters) and an
 * EncryptedPackage stream (the ciphertext of the real docx zip). Both apps
 * write the same MS-OFFCRYPTO formats (Standard = AES-128/SHA-1, Agile =
 * AES-256/SHA-512), and both open either — so we decrypt both and re-encrypt
 * with Agile, matching Word 2013+ defaults.
 */
import {
  decryptOoxml as decryptDocx,
  encryptOoxml as encryptDocx,
  isEncryptedOoxml as isEncryptedDocx,
  OoxmlDecryptError as DocxDecryptError,
  type OoxmlDecryptFailReason as DecryptFailReason,
} from '@genoffice/electron-utils'

export { decryptDocx, encryptDocx, isEncryptedDocx, DocxDecryptError, type DecryptFailReason }

// ── In-memory password state, keyed per renderer + file path.
// Never persisted: a relaunch re-prompts, exactly like Word/WPS. ──

/** Password that can decrypt the bytes currently on disk. */
const diskPasswords = new Map<number, Map<string, string>>()

interface PasswordIntent {
  password: string | null
  revision: number
}

/** Desired password for the next successful save of an existing document. */
const desiredPasswords = new Map<number, Map<string, PasswordIntent>>()
/** Desired password for the next successful save of a pathless document. */
const pendingNewDocPasswords = new Map<number, PasswordIntent>()
let nextIntentRevision = 1

/** paths whose on-disk bytes are currently encrypted (tracked at open/save time:
 *  deterministic, unlike probing the file, which can transiently fail) */
const diskEncryptedPaths = new Map<number, Set<string>>()

function setDiskPassword(wcId: number, filePath: string, password: string | null): void {
  const forWc = diskPasswords.get(wcId) ?? new Map<string, string>()
  if (password) forWc.set(filePath, password)
  else forWc.delete(filePath)
  if (forWc.size > 0) diskPasswords.set(wcId, forWc)
  else diskPasswords.delete(wcId)
}

/** Record the current on-disk password (null when the opened file is plain). */
export function rememberDocPassword(wcId: number, filePath: string, password: string | null): void {
  setDiskPassword(wcId, filePath, password)
}

/** Password for the bytes currently on disk, or null for plain documents. */
export function docPasswordFor(wcId: number, filePath: string | null | undefined): string | null {
  if (!filePath) return null
  return diskPasswords.get(wcId)?.get(filePath) ?? null
}

/**
 * Record the user's desired password for the next save. This deliberately does
 * not alter the password for the current on-disk bytes: recovery copies and
 * reopen must keep using that password until a save actually lands.
 */
export function setDocPassword(
  wcId: number,
  filePath: string | null,
  password: string | null,
): void {
  const intent = { password, revision: nextIntentRevision++ }
  if (filePath) {
    const forWc = desiredPasswords.get(wcId) ?? new Map<string, PasswordIntent>()
    forWc.set(filePath, intent)
    desiredPasswords.set(wcId, forWc)
    return
  }
  pendingNewDocPasswords.set(wcId, intent)
}

/** Capture a cutoff before document-swap cleanup is queued behind an in-flight save. */
export function currentDocPasswordIntentRevision(): number {
  return nextIntentRevision - 1
}

/**
 * The renderer accepted a different document. Discard only intents that
 * existed at the captured cutoff; a password chosen for the replacement
 * document while cleanup waits behind a save must survive.
 */
export function discardDocPasswordIntents(
  wcId: number,
  throughRevision = Number.MAX_SAFE_INTEGER,
): void {
  const forWc = desiredPasswords.get(wcId)
  if (forWc) {
    for (const [path, intent] of forWc) {
      if (intent.revision <= throughRevision) forWc.delete(path)
    }
    if (forWc.size === 0) desiredPasswords.delete(wcId)
  }
  const pending = pendingNewDocPasswords.get(wcId)
  if (pending && pending.revision <= throughRevision) pendingNewDocPasswords.delete(wcId)
}

export interface DocPasswordSaveSnapshot {
  readonly sourcePath: string | null
  readonly password: string | null
  readonly intentRevision: number | null
}

/**
 * Snapshot the effective desired state before a save starts. A failed write
 * leaves the snapshot uncommitted, while a successful write commits exactly
 * these bytes without consuming a newer intent recorded during the await.
 */
export function snapshotDocPassword(
  wcId: number,
  filePath: string | null,
): DocPasswordSaveSnapshot {
  const intent = filePath
    ? desiredPasswords.get(wcId)?.get(filePath)
    : pendingNewDocPasswords.get(wcId)
  return {
    sourcePath: filePath,
    password: intent ? intent.password : docPasswordFor(wcId, filePath),
    intentRevision: intent?.revision ?? null,
  }
}

/**
 * Commit password state only after the atomic file write succeeds. Save As and
 * first-save move a newer concurrent intent to the new path so it remains the
 * desired state for the following save. Returns whether such an intent remains.
 */
export function commitDocPasswordSave(
  wcId: number,
  snapshot: DocPasswordSaveSnapshot,
  savedPath: string,
): boolean {
  setDiskPassword(wcId, savedPath, snapshot.password)
  markDiskEncrypted(wcId, savedPath, snapshot.password !== null)

  const sourcePath = snapshot.sourcePath
  const forWc = desiredPasswords.get(wcId) ?? new Map<string, PasswordIntent>()

  if (sourcePath === null) {
    const current = pendingNewDocPasswords.get(wcId)
    // The target may have stale state from an older document opened in this
    // renderer. Its new baseline is the just-written snapshot.
    forWc.delete(savedPath)
    if (current?.revision === snapshot.intentRevision) {
      pendingNewDocPasswords.delete(wcId)
    } else if (current) {
      pendingNewDocPasswords.delete(wcId)
      forWc.set(savedPath, current)
    }
  } else if (sourcePath === savedPath) {
    const current = forWc.get(sourcePath)
    if (current?.revision === snapshot.intentRevision) forWc.delete(sourcePath)
  } else {
    const current = forWc.get(sourcePath)
    forWc.delete(sourcePath)
    forWc.delete(savedPath)
    if (current && current.revision !== snapshot.intentRevision) {
      forWc.set(savedPath, current)
    }
  }

  if (forWc.size > 0) desiredPasswords.set(wcId, forWc)
  else desiredPasswords.delete(wcId)
  return forWc.has(savedPath)
}

/**
 * Prepare crash-recovery bytes using only the password for the current disk
 * file. Returning null for an encrypted file with missing state is deliberate:
 * skipping a recovery tick is safer than ever writing its plaintext.
 */
export function prepareRecoveryDocx(
  wcId: number,
  filePath: string,
  plainBytes: Buffer,
): Buffer | null {
  if (!isDiskEncrypted(wcId, filePath)) return plainBytes
  const password = docPasswordFor(wcId, filePath)
  return password ? encryptDocx(plainBytes, password) : null
}

/** Decrypt an encrypted recovery copy with the current disk password. */
export async function decryptRecoveryCopy(
  wcId: number,
  filePath: string,
  bytes: Buffer,
): Promise<Buffer> {
  if (!isEncryptedDocx(bytes)) return bytes
  const password = docPasswordFor(wcId, filePath)
  if (!password) {
    throw new DocxDecryptError('unsupported', 'missing password for encrypted recovery copy')
  }
  return decryptDocx(bytes, password)
}

/** File renamed on disk: keep disk and desired state reachable under the new path. */
export function renameDocPassword(wcId: number, oldPath: string, newPath: string): void {
  const forWc = diskPasswords.get(wcId)
  const password = forWc?.get(oldPath)
  if (forWc && password !== undefined) {
    forWc.delete(oldPath)
    forWc.set(newPath, password)
  }
  const desiredForWc = desiredPasswords.get(wcId)
  const desired = desiredForWc?.get(oldPath)
  if (desiredForWc && desired) {
    desiredForWc.delete(oldPath)
    desiredForWc.set(newPath, desired)
  }
  const encSet = diskEncryptedPaths.get(wcId)
  if (encSet?.delete(oldPath)) encSet.add(newPath)
}

/**
 * Record whether the file's on-disk bytes are encrypted; called wherever the
 * main process learns the on-disk state (open, and after every save). Recovery
 * copies key off this: they are only encrypted when the on-disk original is,
 * so a post-crash reopen (which can only ever obtain the on-disk file's
 * password) can always decrypt them.
 */
export function markDiskEncrypted(wcId: number, filePath: string, encrypted: boolean): void {
  const forWc = diskEncryptedPaths.get(wcId) ?? new Set<string>()
  if (encrypted) forWc.add(filePath)
  else forWc.delete(filePath)
  diskEncryptedPaths.set(wcId, forWc)
}

/** The file's on-disk bytes are encrypted (as last opened/saved by this renderer). */
export function isDiskEncrypted(wcId: number, filePath: string): boolean {
  return diskEncryptedPaths.get(wcId)?.has(filePath) ?? false
}

/** Renderer torn down (tab closed): its passwords must not outlive it. */
export function forgetDocPasswords(wcId: number): void {
  diskPasswords.delete(wcId)
  desiredPasswords.delete(wcId)
  pendingNewDocPasswords.delete(wcId)
  diskEncryptedPaths.delete(wcId)
}

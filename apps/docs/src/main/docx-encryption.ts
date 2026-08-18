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
import officeCrypto from 'officecrypto-tool'

const CFB_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
const ENCRYPTED_STREAM_UTF16 = Buffer.from('EncryptedPackage', 'utf16le')

/** CFB (OLE2) container magic — for a .docx path this means "encrypted" (plain docx is a zip) */
export function isCfbFile(bytes: Buffer): boolean {
  return bytes.length >= 8 && bytes.subarray(0, 8).equals(CFB_MAGIC)
}

/** ECMA-376 encrypted OOXML: CFB magic + an EncryptedPackage stream in the directory. */
export function isEncryptedDocx(bytes: Buffer): boolean {
  return isCfbFile(bytes) && bytes.includes(ENCRYPTED_STREAM_UTF16)
}

export type DecryptFailReason = 'wrong-password' | 'unsupported'

export class DocxDecryptError extends Error {
  readonly reason: DecryptFailReason
  constructor(reason: DecryptFailReason, message: string) {
    super(message)
    this.reason = reason
  }
}

/**
 * Decrypt an encrypted docx into plain zip bytes. Throws DocxDecryptError with
 * reason 'wrong-password' (verifier mismatch — reprompt) or 'unsupported'
 * (exotic scheme, e.g. WPS account encryption or Extensible encryption).
 */
export async function decryptDocx(bytes: Buffer, password: string): Promise<Buffer> {
  try {
    return await officeCrypto.decrypt(bytes, { password })
  } catch (err) {
    const message = String((err as Error)?.message ?? err)
    if (message.includes('password is incorrect')) {
      throw new DocxDecryptError('wrong-password', message)
    }
    throw new DocxDecryptError('unsupported', message)
  }
}

/** Re-encrypt plain docx zip bytes with ECMA-376 Agile (Word 2013+ default). */
export function encryptDocx(bytes: Buffer, password: string): Buffer {
  return officeCrypto.encrypt(bytes, { password })
}

// ── In-memory password store, keyed per renderer + file path.
// Never persisted: a relaunch re-prompts, exactly like Word/WPS. ──

const passwords = new Map<number, Map<string, string>>()
/** password chosen for a not-yet-saved document, consumed by its first save */
const pendingNewDocPasswords = new Map<number, string>()
/** paths whose on-disk bytes are currently encrypted (tracked at open/save time:
 *  deterministic, unlike probing the file, which can transiently fail) */
const diskEncryptedPaths = new Map<number, Set<string>>()

export function rememberDocPassword(wcId: number, filePath: string, password: string): void {
  const forWc = passwords.get(wcId) ?? new Map<string, string>()
  forWc.set(filePath, password)
  passwords.set(wcId, forWc)
}

/** Password this renderer opened the file with, or null for plain documents. */
export function docPasswordFor(wcId: number, filePath: string | null | undefined): string | null {
  if (!filePath) return null
  return passwords.get(wcId)?.get(filePath) ?? null
}

/**
 * User set (or cleared, password = null) the document's open password from the
 * ribbon. Pathless documents park it as pending until the first save lands.
 */
export function setDocPassword(
  wcId: number,
  filePath: string | null,
  password: string | null,
): void {
  if (filePath) {
    if (password) rememberDocPassword(wcId, filePath, password)
    else passwords.get(wcId)?.delete(filePath)
    return
  }
  if (password) pendingNewDocPasswords.set(wcId, password)
  else pendingNewDocPasswords.delete(wcId)
}

/**
 * Password for a document's first save (set before it had a path), if any.
 * Read-only: the caller clears it only after the write actually lands, so a
 * failed save keeps the password for the retry.
 */
export function peekPendingDocPassword(wcId: number): string | null {
  return pendingNewDocPasswords.get(wcId) ?? null
}

/** First save landed (or the renderer swapped documents): the pending password is spent/stale. */
export function clearPendingDocPassword(wcId: number): void {
  pendingNewDocPasswords.delete(wcId)
}

/** File renamed on disk: keep this renderer's password reachable under the new path. */
export function renameDocPassword(wcId: number, oldPath: string, newPath: string): void {
  const forWc = passwords.get(wcId)
  const password = forWc?.get(oldPath)
  if (forWc && password !== undefined) {
    forWc.delete(oldPath)
    forWc.set(newPath, password)
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
  passwords.delete(wcId)
  pendingNewDocPasswords.delete(wcId)
  diskEncryptedPaths.delete(wcId)
}

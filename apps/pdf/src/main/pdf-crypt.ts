import { createCipheriv, createDecipheriv, createHash, randomBytes } from 'node:crypto'
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFRawStream,
  PDFRef,
  PDFString,
} from 'pdf-lib'
import type { PdfPermissions } from '../shared/ipc'

/**
 * The standard security handler (PDF 32000 §7.6): password → file key for every
 * revision still found in the wild, enough of it to open password-protected
 * documents for editing and to write AES-256 protection back. Node's crypto
 * supplies MD5/SHA-2/AES; RC4 (revisions 2-3) is short enough to carry here.
 */

/** the 32-byte padding string every legacy password is padded with */
const PAD = new Uint8Array([
  0x28, 0xbf, 0x4e, 0x5e, 0x4e, 0x75, 0x8a, 0x41, 0x64, 0x00, 0x4e, 0x56, 0xff, 0xfa, 0x01, 0x08,
  0x2e, 0x2e, 0x00, 0xb6, 0xd0, 0x68, 0x3e, 0x80, 0x2f, 0x0c, 0xa9, 0xfe, 0x64, 0x53, 0x69, 0x7a,
])

/** the password is not the right one for this document */
export class WrongPasswordError extends Error {
  constructor() {
    super('wrong password')
  }
}

/** UTF-8 without a BOM, cut at the first NUL and at 127 bytes (§7.6.3.1) */
function passwordBytes(password: string): Uint8Array {
  const utf8 = new TextEncoder().encode(password)
  let end = utf8.length
  for (let i = 0; i < utf8.length; i++) {
    if (utf8[i] === 0) {
      end = i
      break
    }
  }
  return utf8.subarray(0, Math.min(end, 127))
}

/** the password padded/truncated to the fixed 32 bytes the legacy algorithms mix */
function paddedPassword(password: string): Uint8Array {
  const bytes = passwordBytes(password)
  const out = new Uint8Array(32)
  out.set(bytes.subarray(0, 32))
  for (let i = bytes.length; i < 32; i++) out[i] = PAD[i - bytes.length]
  return out
}

function md5(...parts: Uint8Array[]): Uint8Array {
  const hash = createHash('md5')
  for (const part of parts) hash.update(part)
  return new Uint8Array(hash.digest())
}

function sha(algorithm: 'sha256' | 'sha384' | 'sha512', data: Uint8Array): Uint8Array {
  return new Uint8Array(createHash(algorithm).update(data).digest())
}

/** RC4 for revisions 2-3; node dropped the cipher, the whole thing is this small */
function rc4(key: Uint8Array, data: Uint8Array): Uint8Array {
  const s = new Uint8Array(256)
  for (let i = 0; i < 256; i++) s[i] = i
  let j = 0
  for (let i = 0; i < 256; i++) {
    j = (j + s[i]! + key[i % key.length]!) & 0xff
    const t = s[i]!
    s[i] = s[j]!
    s[j] = t
  }
  const out = new Uint8Array(data.length)
  let i = 0
  j = 0
  for (let n = 0; n < data.length; n++) {
    i = (i + 1) & 0xff
    j = (j + s[i]!) & 0xff
    const t = s[i]!
    s[i] = s[j]!
    s[j] = t
    out[n] = data[n]! ^ s[(s[i]! + s[j]!) & 0xff]!
  }
  return out
}

function aesCbc(
  mode: 'encrypt' | 'decrypt',
  keyBits: 128 | 256,
  key: Uint8Array,
  iv: Uint8Array,
  data: Uint8Array,
  pad: boolean,
): Uint8Array {
  const name = `aes-${keyBits}-cbc`
  if (mode === 'encrypt') {
    const cipher = createCipheriv(name, key, iv)
    cipher.setAutoPadding(pad)
    return new Uint8Array(Buffer.concat([cipher.update(data), cipher.final()]))
  }
  const decipher = createDecipheriv(name, key, iv)
  decipher.setAutoPadding(false)
  const plain = new Uint8Array(Buffer.concat([decipher.update(data), decipher.final()]))
  if (!pad) return plain
  // PKCS#7: every pad byte equals the count, at least one, never past the buffer
  const count = plain[plain.length - 1] ?? 0
  if (count < 1 || count > 16 || count > plain.length) throw new Error('bad padding')
  for (let i = plain.length - count; i < plain.length; i++) {
    if (plain[i] !== count) throw new Error('bad padding')
  }
  return plain.subarray(0, plain.length - count)
}

function le(value: number, bytes: number): Uint8Array {
  const out = new Uint8Array(bytes)
  for (let i = 0; i < bytes; i++) out[i] = (value >>> (8 * i)) & 0xff
  return out
}

/** Algorithm 2: the file key of a revision 2-4 document */
function legacyFileKey(
  padded: Uint8Array,
  o: Uint8Array,
  p: number,
  id0: Uint8Array,
  keyLength: number,
  revision: number,
  encryptMetadata: boolean,
): Uint8Array {
  const parts: Uint8Array[] = [padded, o.subarray(0, 32), le(p, 4), id0.subarray(0, 16)]
  if (revision >= 4 && !encryptMetadata) parts.push(new Uint8Array([0xff, 0xff, 0xff, 0xff]))
  let hash = md5(...parts)
  if (revision >= 3) {
    for (let i = 0; i < 50; i++) hash = md5(hash.subarray(0, keyLength))
  }
  return hash.subarray(0, keyLength)
}

/** Algorithm 1: the per-object key of a revision ≤ 4 document (AES mixes in the salt) */
function legacyObjectKey(fileKey: Uint8Array, ref: PDFRef, aes: boolean): Uint8Array {
  const parts: Uint8Array[] = [fileKey, le(ref.objectNumber, 3), le(ref.generationNumber, 2)]
  if (aes) parts.push(new Uint8Array([0x73, 0x41, 0x6c, 0x54])) // "sAlT"
  return md5(...parts).subarray(0, Math.min(fileKey.length + 5, 16))
}

/** Algorithm 2.B: the R6 hash. Revision 5 is a plain SHA-256; 6 adds the AES ramp-up loop */
function r6Hash(
  revision: 5 | 6,
  password: Uint8Array,
  salt: Uint8Array,
  uBytes: Uint8Array,
): Uint8Array {
  const input = new Uint8Array(password.length + salt.length + uBytes.length)
  input.set(password, 0)
  input.set(salt, password.length)
  input.set(uBytes, password.length + salt.length)
  let k = sha('sha256', input)
  if (revision === 6) {
    let e: Uint8Array = new Uint8Array([0])
    let i = 0
    while (i < 64 || e[e.length - 1]! > i - 32) {
      const combinedLength = password.length + k.length + uBytes.length
      const combined = new Uint8Array(combinedLength)
      combined.set(password, 0)
      combined.set(k, password.length)
      combined.set(uBytes, password.length + k.length)
      const k1 = new Uint8Array(combinedLength * 64)
      for (let pos = 0; pos < k1.length; pos += combinedLength) k1.set(combined, pos)
      e = aesCbc('encrypt', 128, k.subarray(0, 16), k.subarray(16, 32), k1, false)
      let sum = 0
      for (let b = 0; b < 16; b++) sum += e[b]!
      k = sha(sum % 3 === 0 ? 'sha256' : sum % 3 === 1 ? 'sha384' : 'sha512', e)
      i++
    }
  }
  return k.subarray(0, 32)
}

function aes256DecryptNoPad(key: Uint8Array, data: Uint8Array): Uint8Array {
  return aesCbc('decrypt', 256, key, new Uint8Array(16), data.subarray(0, 32), false)
}

interface LegacyEncrypt {
  revision: 2 | 3 | 4
  fileKey: Uint8Array
  streamCfm: 'V2' | 'AESV2'
  stringCfm: 'V2' | 'AESV2'
}

interface R6Encrypt {
  revision: 5 | 6
  fileKey: Uint8Array
}

/** What a document's /Encrypt dict reduces to once the password opens it */
type Derived = LegacyEncrypt | R6Encrypt

function uintOf(dict: PDFDict, name: string): number {
  const value = dict.get(PDFName.of(name))
  return value instanceof PDFNumber ? value.asNumber() : 0
}

function bytesOf(dict: PDFDict, name: string): Uint8Array {
  // a hex string renders as <...>; the brackets are format, not data
  const hex = String(dict.get(PDFName.of(name)))
    .replace(/^</, '')
    .replace(/>$/, '')
  const out = new Uint8Array(hex.length / 2)
  for (let i = 0; i < out.length; i++) out[i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16)
  return out
}

/**
 * Derive the file key from a password (§7.6.3). Validation happens before any
 * key leaves this function, so a wrong password is a WrongPasswordError and
 * never a silently garbled document.
 */
function deriveKey(encrypt: PDFDict, id0: Uint8Array, password: string): Derived {
  const revision = uintOf(encrypt, 'R')
  const padded = paddedPassword(password)

  if (revision <= 4) {
    const keyLength = Math.min(Math.max(uintOf(encrypt, 'Length'), 40), 128) / 8
    const o = bytesOf(encrypt, 'O')
    const u = bytesOf(encrypt, 'U')
    const p = uintOf(encrypt, 'P')
    const encryptMetadata = readEncryptMetadata(encrypt)
    const cfm = cfmOf(encrypt)
    const compareLength = revision >= 3 ? 16 : u.length
    // the password as the user's (Algorithm 2 + 4/5)
    const userKey = legacyFileKey(padded, o, p, id0, keyLength, revision, encryptMetadata)
    const userExpect = legacyExpectation(userKey, id0, revision)
    if (equal(userExpect.subarray(0, compareLength), u.subarray(0, compareLength))) {
      return { revision: revision as 2 | 3 | 4, fileKey: userKey, streamCfm: cfm, stringCfm: cfm }
    }
    // the password may still be the owner's (Algorithm 7/8): it decrypts /O
    // back into the padded user password, which yields the file key
    const ownerKey = legacyOwnerKey(padded, keyLength, revision)
    const recovered =
      revision >= 3
        ? rc4Rounds(ownerKey, o.subarray(0, 32), 'decrypt')
        : rc4(ownerKey, o.subarray(0, 32))
    const ownerFileKey = legacyFileKey(recovered, o, p, id0, keyLength, revision, encryptMetadata)
    const ownerExpect = legacyExpectation(ownerFileKey, id0, revision)
    if (!equal(ownerExpect.subarray(0, compareLength), u.subarray(0, compareLength))) {
      throw new WrongPasswordError()
    }
    return {
      revision: revision as 2 | 3 | 4,
      fileKey: ownerFileKey,
      streamCfm: cfm,
      stringCfm: cfm,
    }
  }

  const revision6 = revision as 5 | 6
  const u = bytesOf(encrypt, 'U')
  const o = bytesOf(encrypt, 'O')
  const ue = bytesOf(encrypt, 'UE')
  const oe = bytesOf(encrypt, 'OE')
  const pw = passwordBytes(password)
  const userValidation = r6Hash(revision6, pw, u.subarray(32, 40), new Uint8Array(0))
  if (equal(userValidation, u.subarray(0, 32))) {
    // the file key travels in /UE, encrypted under the key-salt hash
    const keyHash = r6Hash(revision6, pw, u.subarray(40, 48), new Uint8Array(0))
    return { revision: revision6, fileKey: aes256DecryptNoPad(keyHash, ue) }
  }
  const ownerValidation = r6Hash(revision6, pw, o.subarray(32, 40), u.subarray(0, 48))
  if (equal(ownerValidation, o.subarray(0, 32))) {
    const keyHash = r6Hash(revision6, pw, o.subarray(40, 48), u.subarray(0, 48))
    return { revision: revision6, fileKey: aes256DecryptNoPad(keyHash, oe) }
  }
  throw new WrongPasswordError()
}

function cfmOf(encrypt: PDFDict): 'V2' | 'AESV2' {
  // revisions ≤ 3 have no crypt filters (RC4); revision 4 reads its StdCF
  const cf = encrypt.get(PDFName.of('CF'))
  if (cf instanceof PDFDict) {
    const std = cf.get(PDFName.of('StdCF'))
    if (std instanceof PDFDict) {
      const cfm = String(std.get(PDFName.of('CFM')))
      if (cfm === '/AESV2') return 'AESV2'
    }
  }
  return 'V2'
}

function readEncryptMetadata(encrypt: PDFDict): boolean {
  const value = encrypt.get(PDFName.of('EncryptMetadata'))
  if (value === undefined) return true
  return String(value) !== 'false'
}

/** Algorithm 3: the RC4 key an /O value was written with (password hash only) */
function legacyOwnerKey(paddedOwner: Uint8Array, keyLength: number, revision: number): Uint8Array {
  let hash = md5(paddedOwner)
  if (revision >= 3) {
    for (let i = 0; i < 50; i++) hash = md5(hash)
  }
  return hash.subarray(0, keyLength)
}

/** rounds 1..19 of key⊕counter RC4, forward to encrypt and 19..0 backward to decrypt */
function rc4Rounds(
  key: Uint8Array,
  data: Uint8Array,
  direction: 'encrypt' | 'decrypt',
): Uint8Array {
  let out = data
  const steps =
    direction === 'encrypt'
      ? [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]
      : [19, 18, 17, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0]
  for (const i of steps) {
    const round = key.map((byte) => byte ^ i)
    out = rc4(round, out)
  }
  return out
}

/**
 * The 32 bytes /U or /O must start with for `key` to be the right one
 * (Algorithm 4/5 for the user entry, 6 for the owner entry). Revision 3+
 * compares only the first 16 — the tail is arbitrary padding.
 */
function legacyExpectation(key: Uint8Array, id0: Uint8Array, revision: number): Uint8Array {
  if (revision <= 2) return rc4(key, PAD)
  const overId = rc4(key, md5(PAD, id0.subarray(0, 16)))
  return rc4Rounds(key, overId, 'encrypt')
}

function equal(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false
  let diff = 0
  for (let i = 0; i < a.length; i++) diff |= a[i]! ^ b[i]!
  return diff === 0
}

function hex(data: Uint8Array): string {
  let out = ''
  for (const byte of data) out += byte.toString(16).padStart(2, '0')
  return out
}

/** Decrypt one string/stream payload with the revision's convention (IV travels with AES data) */
function decryptPayload(
  derived: Derived,
  ref: PDFRef,
  data: Uint8Array,
  isString: boolean,
): Uint8Array {
  if (derived.revision <= 4) {
    const legacy = derived as LegacyEncrypt
    const cfm = isString ? legacy.stringCfm : legacy.streamCfm
    if (cfm === 'AESV2') {
      const key = legacyObjectKey(derived.fileKey, ref, true)
      return aesCbc('decrypt', 128, key, data.subarray(0, 16), data.subarray(16), true)
    }
    const key = legacyObjectKey(derived.fileKey, ref, false)
    return rc4(key, data)
  }
  return aesCbc('decrypt', 256, derived.fileKey, data.subarray(0, 16), data.subarray(16), true)
}

function encryptPayload(derived: R6Encrypt, data: Uint8Array): Uint8Array {
  const iv = randomBytes(16)
  return new Uint8Array([...iv, ...aesCbc('encrypt', 256, derived.fileKey, iv, data, true)])
}

/** /Length of a stream dict must describe the ciphertext the file actually holds */
function replaceStream(
  context: { assign(ref: PDFRef, obj: unknown): void },
  ref: PDFRef,
  stream: PDFRawStream,
  payload: Uint8Array,
): void {
  const dict = stream.dict
  dict.set(PDFName.of('Length'), PDFNumber.of(payload.length))
  context.assign(ref, PDFRawStream.of(dict, payload))
}

/** Every string and stream of the document, except the /Encrypt dict's own */
function transformPayloads(
  doc: PDFDocument,
  skip: PDFRef | null,
  transform: (ref: PDFRef, data: Uint8Array, isString: boolean) => Uint8Array,
): void {
  const context = doc.context
  for (const [ref, obj] of context.enumerateIndirectObjects()) {
    if (skip && ref.toString() === skip.toString()) continue
    if (obj instanceof PDFRawStream) {
      replaceStream(context, ref, obj, transform(ref, obj.contents, false))
      continue
    }
    const walk = (node: unknown): void => {
      if (node instanceof PDFDict) {
        for (const [key, value] of node.entries()) {
          if (value instanceof PDFString || value instanceof PDFHexString) {
            const plain = transform(ref, value.asBytes(), true)
            node.set(key, PDFHexString.of(hex(plain)))
          } else {
            walk(value)
          }
        }
      } else if (node instanceof PDFArray) {
        for (let i = 0; i < node.size(); i++) {
          const value = node.get(i)
          if (value instanceof PDFString || value instanceof PDFHexString) {
            const plain = transform(ref, value.asBytes(), true)
            if (plain !== value.asBytes()) node.set(i, PDFHexString.of(hex(plain)))
          } else {
            walk(value)
          }
        }
      }
    }
    walk(obj)
  }
}

/** active-low bits; a denied permission clears its bit (§7.6.3.2 Table 24) */
export function permissionsToP(permissions: PdfPermissions): number {
  let p = -1
  const clear = (mask: number): void => {
    p &= ~mask
  }
  if (!permissions.printing) clear(0x4 | 0x800)
  if (!permissions.modifying) clear(0x8 | 0x400)
  if (!permissions.annotating) clear(0x20 | 0x100)
  if (!permissions.copying) clear(0x10)
  return p | 0
}

export function permissionsFromP(p: number): PdfPermissions {
  return {
    printing: (p & 0x4) !== 0 || (p & 0x800) !== 0,
    modifying: (p & 0x8) !== 0 || (p & 0x400) !== 0,
    annotating: (p & 0x20) !== 0 || (p & 0x100) !== 0,
    copying: (p & 0x10) !== 0,
  }
}

/** The document's first /ID element as raw bytes ('' when the file has none) */
function documentId0(doc: PDFDocument): Uint8Array {
  const id = doc.context.trailerInfo.ID
  if (!(id instanceof PDFArray) || id.size() === 0) return new Uint8Array(0)
  const hex0 = String(id.get(0)).replace(/^</, '').replace(/>$/, '')
  const out = new Uint8Array(hex0.length / 2)
  for (let i = 0; i < out.length; i++) out[i] = parseInt(hex0.slice(i * 2, i * 2 + 2), 16)
  return out
}

export interface DecryptedPdf {
  bytes: Uint8Array
  /** the /P of the removed protection, so the session can offer it back */
  permissions: PdfPermissions
  /** whether a password was actually required to open (false: owner-only file) */
  neededPassword: boolean
}

/**
 * Open a password-protected document for editing: every string and stream is
 * decrypted in place, the /Encrypt dictionary is dropped, and plain bytes come
 * back. The whole rest of the app — markup, text edits, forms, save, signing —
 * then works on the document as if it had never been encrypted.
 */
export async function decryptDocument(bytes: Uint8Array, password: string): Promise<DecryptedPdf> {
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false })
  const encryptRef = doc.context.trailerInfo.Encrypt
  const encryptDict = encryptRef ? doc.context.lookup(encryptRef) : undefined
  if (!(encryptDict instanceof PDFDict)) {
    return {
      bytes,
      permissions: permissionsFromP(-1),
      neededPassword: false,
    }
  }
  const derived = deriveKey(encryptDict, documentId0(doc), password)
  transformPayloads(doc, encryptRef as PDFRef, (ref, data) =>
    decryptPayload(derived, ref, data, true),
  )
  delete (doc.context.trailerInfo as { Encrypt?: unknown }).Encrypt
  return {
    bytes: await doc.save({ useObjectStreams: false }),
    permissions: permissionsFromP(uintOf(encryptDict, 'P')),
    neededPassword: password.length > 0,
  }
}

export interface PdfProtection {
  userPassword: string
  /** defaults to the user password; the owner can always open and change limits */
  ownerPassword?: string
  permissions: PdfPermissions
}

/**
 * Write AES-256 protection into a document that save() is about to serialize:
 * strings and streams are encrypted in place and the trailer gains /Encrypt.
 * The handler is written as V=5 (R6) — the pairing every mainstream producer
 * and pdf.js itself reads; ISO 32000-2's V=6 spelling is not widely accepted.
 * Save the document with `useObjectStreams: false` — object and xref streams
 * are created during serialization, after this pass could reach them.
 */
export function encryptDocument(doc: PDFDocument, protection: PdfProtection): void {
  const context = doc.context
  const pw = passwordBytes(protection.userPassword)
  const validationSalt = randomBytes(8)
  const keySalt = randomBytes(8)
  const userHash = r6Hash(6, pw, validationSalt, new Uint8Array(0))
  const keyHash = r6Hash(6, pw, keySalt, new Uint8Array(0))
  const fileKey = randomBytes(32)
  const u = new Uint8Array([...userHash, ...validationSalt, ...keySalt])
  // the file key travels in /UE under the key-salt hash
  const ue = aesCbc('encrypt', 256, keyHash, new Uint8Array(16), fileKey, false)

  const owner = protection.ownerPassword ?? protection.userPassword
  const ownerBytes = passwordBytes(owner)
  const ownerValidationSalt = randomBytes(8)
  const ownerKeySalt = randomBytes(8)
  const ownerHash = r6Hash(6, ownerBytes, ownerValidationSalt, u)
  const ownerKeyHash = r6Hash(6, ownerBytes, ownerKeySalt, u)
  const o = new Uint8Array([...ownerHash, ...ownerValidationSalt, ...ownerKeySalt])
  const oe = aesCbc('encrypt', 256, ownerKeyHash, new Uint8Array(16), fileKey, false)
  // /Perms: the permission bits, signed so readers can trust them (§7.6.4.4):
  // P as 32-bit little-endian, four required 0xFF bytes, the metadata flag as
  // 'T'/'F', the 'adb' marker, four throwaway bytes
  const p = permissionsToP(protection.permissions)
  const permsData = new Uint8Array(16)
  permsData.set(le(p, 4), 0)
  permsData.set([0xff, 0xff, 0xff, 0xff, 0x54, 0x61, 0x64, 0x62], 4)
  permsData.set(randomBytes(4), 12)
  const ecb = createCipheriv('aes-256-ecb', fileKey, null)
  ecb.setAutoPadding(false)
  const perms = new Uint8Array(Buffer.concat([ecb.update(permsData), ecb.final()]))

  const encrypt = context.obj({
    Filter: 'Standard',
    V: 5,
    R: 6,
    Length: 256,
    P: PDFNumber.of(permissionsToP(protection.permissions)),
    O: PDFHexString.of(hex(o)),
    U: PDFHexString.of(hex(u)),
    OE: PDFHexString.of(hex(oe)),
    UE: PDFHexString.of(hex(ue)),
    Perms: PDFHexString.of(hex(perms)),
  })
  const cf = context.obj({
    StdCF: context.obj({ CFM: 'AESV3', AuthEvent: 'DocOpen', Length: 32 }),
  })
  encrypt.set(PDFName.of('CF'), cf)
  encrypt.set(PDFName.of('StmF'), PDFName.of('StdCF'))
  encrypt.set(PDFName.of('StrF'), PDFName.of('StdCF'))
  encrypt.set(PDFName.of('EncryptMetadata'), context.obj(true))
  const encryptRef = context.nextRef()
  context.assign(encryptRef, encrypt)
  context.trailerInfo.Encrypt = encryptRef

  if (!context.trailerInfo.ID) {
    const id = randomBytes(16).toString('hex').toUpperCase()
    context.trailerInfo.ID = context.obj([PDFHexString.of(id), PDFHexString.of(id)])
  }

  const derived: R6Encrypt = { revision: 6, fileKey }
  transformPayloads(doc, encryptRef, (_ref, data) => encryptPayload(derived, data))
}

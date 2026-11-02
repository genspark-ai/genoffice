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
  PDFStream,
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

/** a certificate (public-key) encrypted document: there is no password to ask for */
export class CertificateEncryptionError extends Error {
  constructor() {
    super('certificate encryption')
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
  streamCfm: 'V2' | 'AESV2' | 'Identity'
  stringCfm: 'V2' | 'AESV2' | 'Identity'
}

interface R6Encrypt {
  revision: 5 | 6
  fileKey: Uint8Array
  /** /StmF or /StrF naming /Identity: that class travels in the clear */
  stmIdentity: boolean
  strIdentity: boolean
}

/** What a document's /Encrypt dict reduces to once the password opens it */
type Derived = LegacyEncrypt | R6Encrypt

/** which password opened the document — the owner's ignores the /P limits */
export type UnlockedAs = 'user' | 'owner'

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
function deriveKey(
  encrypt: PDFDict,
  id0: Uint8Array,
  password: string,
): { derived: Derived; unlockedAs: UnlockedAs } {
  const revision = uintOf(encrypt, 'R')
  const padded = paddedPassword(password)

  if (revision <= 4) {
    // /Length is optional: V4 files take it from the crypt filter (bytes), older ones mean 40
    const declared = uintOf(encrypt, 'Length') || (revision === 4 ? filterKeyBits(encrypt) : 40)
    const keyLength = Math.min(Math.max(declared, 40), 128) / 8
    const o = bytesOf(encrypt, 'O')
    const u = bytesOf(encrypt, 'U')
    const p = uintOf(encrypt, 'P')
    const encryptMetadata = readEncryptMetadata(encrypt)
    const cfms = cfmsOf(encrypt)
    const compareLength = revision >= 3 ? 16 : u.length
    // the password as the user's (Algorithm 2 + 4/5)
    const userKey = legacyFileKey(padded, o, p, id0, keyLength, revision, encryptMetadata)
    const userExpect = legacyExpectation(userKey, id0, revision)
    if (equal(userExpect.subarray(0, compareLength), u.subarray(0, compareLength))) {
      return {
        derived: {
          revision: revision as 2 | 3 | 4,
          fileKey: userKey,
          streamCfm: cfms.stream,
          stringCfm: cfms.string,
        },
        unlockedAs: 'user',
      }
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
      derived: {
        revision: revision as 2 | 3 | 4,
        fileKey: ownerFileKey,
        streamCfm: cfms.stream,
        stringCfm: cfms.string,
      },
      unlockedAs: 'owner',
    }
  }

  const revision6 = revision as 5 | 6
  const u = bytesOf(encrypt, 'U')
  const o = bytesOf(encrypt, 'O')
  const ue = bytesOf(encrypt, 'UE')
  const oe = bytesOf(encrypt, 'OE')
  const pw = passwordBytes(password)
  const cfms = cfmsOf(encrypt)
  const userValidation = r6Hash(revision6, pw, u.subarray(32, 40), new Uint8Array(0))
  if (equal(userValidation, u.subarray(0, 32))) {
    // the file key travels in /UE, encrypted under the key-salt hash
    const keyHash = r6Hash(revision6, pw, u.subarray(40, 48), new Uint8Array(0))
    return {
      derived: {
        revision: revision6,
        fileKey: aes256DecryptNoPad(keyHash, ue),
        stmIdentity: cfms.stream === 'Identity',
        strIdentity: cfms.string === 'Identity',
      },
      unlockedAs: 'user',
    }
  }
  const ownerValidation = r6Hash(revision6, pw, o.subarray(32, 40), u.subarray(0, 48))
  if (equal(ownerValidation, o.subarray(0, 32))) {
    const keyHash = r6Hash(revision6, pw, o.subarray(40, 48), u.subarray(0, 48))
    return {
      derived: {
        revision: revision6,
        fileKey: aes256DecryptNoPad(keyHash, oe),
        stmIdentity: cfms.stream === 'Identity',
        strIdentity: cfms.string === 'Identity',
      },
      unlockedAs: 'owner',
    }
  }
  throw new WrongPasswordError()
}

/** key bits from the stream crypt filter's /Length (bytes; a few writers put bits) */
function filterKeyBits(encrypt: PDFDict): number {
  const stmF = encrypt.get(PDFName.of('StmF'))
  const filterName = stmF instanceof PDFName ? stmF.asString().slice(1) : 'StdCF'
  const cf = encrypt.get(PDFName.of('CF'))
  const filter = cf instanceof PDFDict ? cf.get(PDFName.of(filterName)) : undefined
  const length = filter instanceof PDFDict ? uintOf(filter, 'Length') : 0
  if (length === 0) return 128
  return length >= 40 ? length : length * 8
}

/** the cipher each payload class runs through, per /StmF and /StrF (§7.6.3.2) */
function cfmsOf(encrypt: PDFDict): {
  stream: 'V2' | 'AESV2' | 'Identity'
  string: 'V2' | 'AESV2' | 'Identity'
} {
  // /StmF and /StrF name the crypt filter their class runs through; a name of
  // /Identity leaves that class in the clear, and the default is StdCF
  const of = (which: 'StmF' | 'StrF'): 'V2' | 'AESV2' | 'Identity' => {
    const name = encrypt.get(PDFName.of(which))
    const filterName = name instanceof PDFName ? name.asString().slice(1) : 'StdCF'
    if (filterName === 'Identity') return 'Identity'
    const cf = encrypt.get(PDFName.of('CF'))
    if (cf instanceof PDFDict) {
      const filter = cf.get(PDFName.of(filterName))
      if (filter instanceof PDFDict) {
        const cfm = String(filter.get(PDFName.of('CFM')))
        if (cfm === '/AESV2') return 'AESV2'
        if (cfm === '/Identity') return 'Identity'
      }
    }
    return 'V2'
  }
  return { stream: of('StmF'), string: of('StrF') }
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
  // an empty payload was empty plaintext: AES-CBC would have nothing to unpick
  if (data.length === 0) return data
  if (derived.revision <= 4) {
    const legacy = derived as LegacyEncrypt
    const cfm = isString ? legacy.stringCfm : legacy.streamCfm
    if (cfm === 'Identity') return data
    if (cfm === 'AESV2') {
      const key = legacyObjectKey(derived.fileKey, ref, true)
      return aesCbc('decrypt', 128, key, data.subarray(0, 16), data.subarray(16), true)
    }
    const key = legacyObjectKey(derived.fileKey, ref, false)
    return rc4(key, data)
  }
  const r6 = derived as R6Encrypt
  if (isString ? r6.strIdentity : r6.stmIdentity) return data
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
  stream: PDFStream,
  payload: Uint8Array,
): void {
  const dict = stream.dict
  dict.set(PDFName.of('Length'), PDFNumber.of(payload.length))
  context.assign(ref, PDFRawStream.of(dict, payload))
}

/** /Type /Metadata: exempt from encryption when the document says EncryptMetadata false */
function isMetadataStream(stream: PDFStream): boolean {
  return String(stream.dict.get(PDFName.of('Type'))) === '/Metadata'
}

/**
 * Encrypt (or decrypt) every string reachable from a dictionary or array, at any
 * depth. Strings in a stream's dictionary travel with the object like any other
 * string of it (§7.6.3.2) — embedded-file /Params entries, /ModDate and friends —
 * so a stream's dict is walked the same way as a plain dictionary.
 */
function replaceStrings(node: unknown, transform: (data: Uint8Array) => Uint8Array): void {
  const walk = (current: unknown): void => {
    if (current instanceof PDFDict) {
      for (const [key, value] of current.entries()) {
        if (value instanceof PDFString || value instanceof PDFHexString) {
          current.set(key, PDFHexString.of(hex(transform(value.asBytes()))))
        } else {
          walk(value)
        }
      }
    } else if (current instanceof PDFArray) {
      for (let i = 0; i < current.size(); i++) {
        const value = current.get(i)
        if (value instanceof PDFString || value instanceof PDFHexString) {
          current.set(i, PDFHexString.of(hex(transform(value.asBytes()))))
        } else {
          walk(value)
        }
      }
    }
  }
  walk(node)
}

/** Every string and stream of the document, except the /Encrypt dict's own */
function transformPayloads(
  doc: PDFDocument,
  skip: PDFRef | null,
  transform: (ref: PDFRef, data: Uint8Array, isString: boolean) => Uint8Array,
  /** false when the source's /EncryptMetadata exempts metadata streams */
  encryptMetadata = true,
): void {
  const context = doc.context
  for (const [ref, obj] of context.enumerateIndirectObjects()) {
    if (skip && ref.toString() === skip.toString()) continue
    const strings = (data: Uint8Array): Uint8Array => transform(ref, data, true)
    if (obj instanceof PDFStream) {
      if (!encryptMetadata && isMetadataStream(obj)) continue
      // pdf-lib's own content streams are PDFFlateStream, not PDFRawStream: only
      // raw streams were encrypted, so every rebuilt page's content stream was
      // written in plaintext inside a file declared encrypted — a corrupt save
      replaceStream(context, ref, obj, transform(ref, obj.getContents(), false))
      // the stream's dictionary strings are object strings too (/Params /ModDate)
      replaceStrings(obj.dict, strings)
      continue
    }
    // a top-level indirect string object (a string that *is* the object) is not
    // reachable from a dictionary walk
    if (obj instanceof PDFString || obj instanceof PDFHexString) {
      context.assign(ref, PDFHexString.of(hex(strings(obj.asBytes()))))
      continue
    }
    replaceStrings(obj, strings)
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

// ─── Byte-level pre-pass (object / cross-reference streams) ─────────────────
/**
 * pdf-lib parses the cross-reference data — and decompresses every /ObjStm —
 * before any decryption here runs, so a real-world file (object streams, xref
 * streams) throws inside the parser and never reaches the per-object pass. The
 * streams are therefore decrypted on the raw bytes first: the cross-reference
 * stream is never encrypted (§7.5.8.2), and an /ObjStm is encrypted as a whole
 * (§7.5.7), so decrypting its payload is exactly what lets the parser read the
 * objects inside. Strings of top-level objects stay for the per-object pass;
 * objects inside an /ObjStm are not encrypted individually, so their strings
 * must not be touched there.
 */

interface RawObject {
  num: number
  gen: number
  dictStart: number
  dictEnd: number
  /** stream payload range; dataStart === -1 when the object has no stream */
  dataStart: number
  dataEnd: number
}

function hexToBytes(value: string): Uint8Array {
  const clean = value.replace(/\s+/g, '')
  const out = new Uint8Array(Math.ceil(clean.length / 2))
  for (let i = 0; i < out.length; i++) {
    out[i] = parseInt(clean.slice(i * 2, i * 2 + 2).padEnd(2, '0'), 16)
  }
  return out
}

/** every top-level `N G obj`, with its dictionary and stream ranges */
function scanRawObjects(text: string): RawObject[] {
  const headers: { index: number; num: number; gen: number; dictStart: number }[] = []
  const header = /(?:^|[\r\n\s])(\d+)\s+(\d+)\s+obj\b/g
  let match: RegExpExecArray | null
  while ((match = header.exec(text)) !== null) {
    headers.push({
      index: match.index,
      num: Number(match[1]),
      gen: Number(match[2]),
      dictStart: header.lastIndex,
    })
  }
  const out: RawObject[] = []
  for (const [i, start] of headers.entries()) {
    const limit = headers[i + 1]?.index ?? text.length
    const slice = text.slice(start.dictStart, limit)
    // the first `stream` after the header is the keyword: the payload starts
    // past it, so a payload's own bytes cannot win the race
    const streamAt = /\bstream(\r\n|\n|\r)/.exec(slice)
    if (streamAt === null) {
      const endobj = slice.indexOf('endobj')
      const dictEnd = start.dictStart + (endobj === -1 ? slice.length : endobj)
      out.push({
        num: start.num,
        gen: start.gen,
        dictStart: start.dictStart,
        dictEnd,
        dataStart: -1,
        dataEnd: -1,
      })
      continue
    }
    const dictEnd = start.dictStart + streamAt.index
    const dataStart = start.dictStart + streamAt.index + streamAt[0].length
    const dict = text.slice(start.dictStart, dictEnd)
    const indirect = /\/Length\s+(\d+)\s+(\d+)\s+R/.exec(dict)
    const direct = indirect === null ? /\/Length\s+(\d+)/.exec(dict) : null
    let dataEnd = -1
    if (direct !== null) {
      const candidate = dataStart + Number(direct[1])
      if (/^\s*endstream/.test(text.slice(candidate, candidate + 24))) dataEnd = candidate
    }
    if (dataEnd === -1) {
      // no usable /Length (indirect, or a producer that got it wrong): the
      // keyword bounds the payload; a producer's own EOL is not part of it
      const endstream = text.indexOf('endstream', dataStart)
      if (endstream === -1) {
        out.push({
          num: start.num,
          gen: start.gen,
          dictStart: start.dictStart,
          dictEnd,
          dataStart: -1,
          dataEnd: -1,
        })
        continue
      }
      dataEnd = endstream
      if (text.slice(dataEnd - 2, dataEnd) === '\r\n') dataEnd -= 2
      else if (text[dataEnd - 1] === '\n' || text[dataEnd - 1] === '\r') dataEnd -= 1
    }
    out.push({
      num: start.num,
      gen: start.gen,
      dictStart: start.dictStart,
      dictEnd,
      dataStart,
      dataEnd,
    })
  }
  return out
}

/** the /Encrypt dictionary as the file stores it (it is never encrypted itself) */
function rawEncryptDict(text: string): { num: number; gen: number; body: string } | null {
  const ref = /\/Encrypt\s+(\d+)\s+(\d+)\s+R/.exec(text)
  if (ref === null) return null
  const num = Number(ref[1])
  const gen = Number(ref[2])
  const found = scanRawObjects(text).find((o) => o.num === num && o.gen === gen)
  if (found === undefined) return null
  return { num, gen, body: text.slice(found.dictStart, found.dictEnd) }
}

function rawId0(text: string): Uint8Array {
  const id = /\/ID\s*\[\s*<([0-9a-fA-F\s]+)>/.exec(text)
  return id === null ? new Uint8Array(0) : hexToBytes(id[1]!)
}

/** the trailer /ID entries; a legacy file key (revision 4 and down) is derived from ID[0] */
function rawFileId(text: string): Uint8Array[] {
  const id = /\/ID\s*\[\s*<([0-9a-fA-F\s]*)>\s*<([0-9a-fA-F\s]*)>\s*\]/.exec(text)
  if (id === null) return []
  return [hexToBytes(id[1]!), hexToBytes(id[2]!)]
}

/** the /Encrypt dictionary reduced to the fields key derivation and re-emission need */
export interface EncryptEntries {
  filter: string
  v: number
  r: number
  length: number
  p: number
  o: Uint8Array
  u: Uint8Array
  oe: Uint8Array
  ue: Uint8Array
  perms: Uint8Array
  /** named crypt filters from /CF; /StmF and /StrF name the ones in use */
  cf: Record<string, { cfm: string; authEvent?: string; lengthBytes?: number }>
  stmF: string
  strF: string
  encryptMetadata: boolean
}

/** Everything a save needs to protect new bytes under the source file's own protection. */
export interface OriginalProtection {
  derived: Derived
  entries: EncryptEntries
  /** trailer /ID entries: a legacy file key is derived from ID[0] too */
  fileId: Uint8Array[]
}

/** a hex "<...>" or literal "(...)" entry; producers write /O and /U both ways */
function entryBytesOf(body: string, name: string): Uint8Array {
  const hexMatch = new RegExp('/' + name + '\\s*<([0-9a-fA-F\\s]+)>').exec(body)
  if (hexMatch !== null) return hexToBytes(hexMatch[1]!)
  const literal = new RegExp('/' + name + '\\s*\\(((?:[^()\\\\]|\\\\.)*)\\)').exec(body)
  if (literal === null) return new Uint8Array(0)
  const raw = literal[1]!
  const out: number[] = []
  for (let i = 0; i < raw.length; i++) {
    if (raw[i] !== '\\') {
      out.push(raw.charCodeAt(i) & 0xff)
      continue
    }
    const esc = raw[++i]
    if (esc === undefined) break
    if (esc === 'n') out.push(0x0a)
    else if (esc === 'r') out.push(0x0d)
    else if (esc === 't') out.push(0x09)
    else if (esc === 'b') out.push(0x08)
    else if (esc === 'f') out.push(0x0c)
    else if (esc >= '0' && esc <= '7') {
      let octal = esc
      while (octal.length < 3 && raw[i + 1] >= '0' && raw[i + 1] <= '7') octal += raw[++i]
      out.push(parseInt(octal, 8) & 0xff)
    } else out.push(esc.charCodeAt(0) & 0xff) // escaped parentheses and backslash: itself
  }
  return new Uint8Array(out)
}

/** the raw /Encrypt dictionary body, reduced to what derivation and a re-emitting writer share */
function readEncryptEntries(body: string): EncryptEntries {
  const number = (name: string): number => {
    const all = [...body.matchAll(new RegExp('/' + name + '\\s+(-?\\d+)', 'g'))].map((m) =>
      Number(m[1]),
    )
    if (all.length === 0) return 0
    // a crypt filter's own /Length counts bytes (16/32) while the root /Length
    // counts bits (40/128/256): the larger value is the root's
    return name === 'Length' ? Math.max(...all) : all[0]!
  }
  // a crypt filter block; /StmF and /StrF name the one each payload class uses
  const filters: Record<string, { cfm: string; authEvent?: string; lengthBytes?: number }> = {}
  for (const match of body.matchAll(/\/(\w+)\s*<<([^<>]*)>>/g)) {
    const inner = match[2] ?? ''
    const authEvent = /\/AuthEvent\s*\/(\w+)/.exec(inner)?.[1]
    const lengthBytes = /\/Length\s+(\d+)/.exec(inner)?.[1]
    filters[match[1]!] = {
      cfm: /\/CFM\s*\/(\w+)/.exec(inner)?.[1] ?? 'V2',
      ...(authEvent === undefined ? {} : { authEvent }),
      ...(lengthBytes === undefined ? {} : { lengthBytes: Number(lengthBytes) }),
    }
  }
  return {
    filter: /\/Filter\s*\/(\w+)/.exec(body)?.[1] ?? 'Standard',
    v: number('V'),
    r: number('R'),
    length: number('Length'),
    p: number('P'),
    o: entryBytesOf(body, 'O'),
    u: entryBytesOf(body, 'U'),
    oe: entryBytesOf(body, 'OE'),
    ue: entryBytesOf(body, 'UE'),
    perms: entryBytesOf(body, 'Perms'),
    cf: filters,
    stmF: /\/StmF\s*\/(\w+)/.exec(body)?.[1] ?? 'StdCF',
    strF: /\/StrF\s*\/(\w+)/.exec(body)?.[1] ?? 'StdCF',
    encryptMetadata: !/\/EncryptMetadata\s+false/.test(body),
  }
}

/** the /Encrypt dictionary to write; a carried protection re-emits the source's entries */
function encryptDictFor(context: PDFDocument['context'], entries: EncryptEntries): PDFDict {
  const filters: Record<string, PDFDict> = {}
  for (const [name, filter] of Object.entries(entries.cf)) {
    filters[name] = context.obj({
      CFM: filter.cfm,
      ...(filter.authEvent === undefined ? {} : { AuthEvent: filter.authEvent }),
      ...(filter.lengthBytes === undefined ? {} : { Length: filter.lengthBytes }),
    }) as PDFDict
  }
  const dict = context.obj({
    Filter: entries.filter,
    V: entries.v,
    R: entries.r,
    ...(entries.length >= 40 ? { Length: entries.length } : {}),
    P: entries.p,
    O: PDFHexString.of(hex(entries.o)),
    U: PDFHexString.of(hex(entries.u)),
    ...(entries.oe.length > 0 ? { OE: PDFHexString.of(hex(entries.oe)) } : {}),
    ...(entries.ue.length > 0 ? { UE: PDFHexString.of(hex(entries.ue)) } : {}),
    ...(entries.perms.length > 0 ? { Perms: PDFHexString.of(hex(entries.perms)) } : {}),
    ...(Object.keys(filters).length > 0 ? { CF: context.obj(filters) } : {}),
    StmF: entries.stmF,
    StrF: entries.strF,
    ...(entries.encryptMetadata ? {} : { EncryptMetadata: false }),
  })
  if (!(dict instanceof PDFDict)) throw new Error('could not build the encryption dictionary')
  return dict
}

/** the raw /Encrypt entries as a PDFDict, so deriveKey can read them */
async function deriveKeyFromEntries(
  entries: EncryptEntries,
  id0: Uint8Array,
  password: string,
): Promise<{ derived: Derived; unlockedAs: UnlockedAs }> {
  const scratch = await PDFDocument.create()
  return deriveKey(encryptDictFor(scratch.context, entries), id0, password)
}

/** decrypt every top-level stream payload in place; returns the new bytes */
function predecryptStreams(
  bytes: Uint8Array,
  derived: Derived,
  encryptMetadata: boolean,
): { bytes: Uint8Array; topLevel: Set<number> } {
  const text = Buffer.from(bytes).toString('latin1')
  const objects = scanRawObjects(text)
  const topLevel = new Set(objects.map((o) => o.num))
  const edits: { start: number; end: number; replacement: Uint8Array }[] = []
  const indirectLengths = new Map<number, number>()
  for (const obj of objects) {
    if (obj.dataStart < 0) continue
    const dict = text.slice(obj.dictStart, obj.dictEnd)
    // the cross-reference stream is never encrypted, and a metadata stream is
    // exempt when the document says so
    if (/\/Type\s*\/XRef\b/.test(dict)) continue
    if (!encryptMetadata && /\/Type\s*\/Metadata\b/.test(dict)) continue
    const payload = bytes.subarray(obj.dataStart, obj.dataEnd)
    let plain: Uint8Array
    try {
      plain = decryptPayload(derived, PDFRef.of(obj.num, obj.gen), payload, false)
    } catch (err) {
      throw new Error(`could not decrypt stream object ${obj.num}: ${String(err)}`, { cause: err })
    }
    edits.push({ start: obj.dataStart, end: obj.dataEnd, replacement: new Uint8Array(plain) })
    const indirect = /\/Length\s+(\d+)\s+(\d+)\s+R/.exec(dict)
    if (indirect !== null) {
      indirectLengths.set(Number(indirect[1]), plain.length)
      continue
    }
    const direct = /\/Length\s+(\d+)/.exec(dict)
    if (direct !== null) {
      const at = obj.dictStart + direct.index + direct[0].indexOf(direct[1]!)
      edits.push({
        start: at,
        end: at + direct[1]!.length,
        replacement: new TextEncoder().encode(String(plain.length)),
      })
    }
  }
  // an indirect /Length lives in an object of its own; its value is the only
  // thing there, so the body is replaced wholesale
  for (const obj of objects) {
    const next = indirectLengths.get(obj.num)
    if (next === undefined || obj.dataStart >= 0) continue
    const raw = text.slice(obj.dictStart, obj.dictEnd)
    const body = raw.trim()
    if (!/^-?\d+$/.test(body)) continue
    const at = obj.dictStart + raw.indexOf(body)
    edits.push({
      start: at,
      end: at + body.length,
      replacement: new TextEncoder().encode(String(next)),
    })
  }
  if (edits.length === 0) return { bytes, topLevel }
  edits.sort((a, b) => a.start - b.start)
  // a plain Uint8Array, not a Node Buffer: pdf-lib's instance check runs in the
  // renderer's realm under jsdom, where a Buffer is not one
  let total = 0
  let cursor = 0
  for (const edit of edits) {
    total += edit.start - cursor + edit.replacement.length
    cursor = edit.end
  }
  total += bytes.length - cursor
  const out = new Uint8Array(total)
  let at = 0
  cursor = 0
  for (const edit of edits) {
    out.set(bytes.subarray(cursor, edit.start), at)
    at += edit.start - cursor
    out.set(edit.replacement, at)
    at += edit.replacement.length
    cursor = edit.end
  }
  out.set(bytes.subarray(cursor), at)
  return { bytes: out, topLevel }
}

/** decrypt the strings of top-level objects (ObjStm residents are already plain) */
function decryptStringsOf(
  doc: PDFDocument,
  skip: PDFRef | null,
  topLevel: Set<number>,
  transform: (ref: PDFRef, data: Uint8Array) => Uint8Array,
  encryptMetadata: boolean,
): void {
  const context = doc.context
  for (const [ref, obj] of context.enumerateIndirectObjects()) {
    if (skip !== null && ref.toString() === skip.toString()) continue
    if (!topLevel.has(ref.objectNumber)) continue
    const strings = (data: Uint8Array): Uint8Array => transform(ref, data)
    // a stream's payload was decrypted on the raw bytes already; only its
    // dictionary strings are left to this pass
    if (obj instanceof PDFStream) {
      if (!encryptMetadata && isMetadataStream(obj)) continue
      replaceStrings(obj.dict, strings)
      continue
    }
    // a top-level indirect string object (the string is the whole object)
    if (obj instanceof PDFString || obj instanceof PDFHexString) {
      context.assign(ref, PDFHexString.of(hex(strings(obj.asBytes()))))
      continue
    }
    replaceStrings(obj, strings)
  }
}

export interface DecryptedPdf {
  bytes: Uint8Array
  /** the /P of the removed protection, so the session can offer it back */
  permissions: PdfPermissions
  /** whether a password was actually required to open (false: owner-only file) */
  neededPassword: boolean
  /** which password opened it; the owner's ignores the /P limits */
  unlockedAs: UnlockedAs
  /** the source file's own protection; absent when the source was plain */
  protection?: OriginalProtection
}

/**
 * Open a password-protected document for editing: every string and stream is
 * decrypted in place, the /Encrypt dictionary is dropped, and plain bytes come
 * back. The whole rest of the app — markup, text edits, forms, save, signing —
 * then works on the document as if it had never been encrypted.
 */
export async function decryptDocument(bytes: Uint8Array, password: string): Promise<DecryptedPdf> {
  const text = Buffer.from(bytes).toString('latin1')
  const encrypt = rawEncryptDict(text)
  if (encrypt === null) {
    return { bytes, permissions: permissionsFromP(-1), neededPassword: false, unlockedAs: 'user' }
  }
  // a certificate-encrypted document carries /Recipients envelopes instead of
  // the password entries: there is no password to type, and saying so is the
  // whole answer this build can give
  if (/\/Recipients\b/.test(encrypt.body) || /\/Filter\s*\/(?!Standard\b)/.test(encrypt.body)) {
    throw new CertificateEncryptionError()
  }
  const entries = readEncryptEntries(encrypt.body)
  const { derived, unlockedAs } = await deriveKeyFromEntries(entries, rawId0(text), password)
  // streams first, on the raw bytes: the parser cannot read an encrypted
  // /ObjStm, and nothing below this line runs until it can
  const pre = predecryptStreams(bytes, derived, entries.encryptMetadata)
  const doc = await PDFDocument.load(pre.bytes, { ignoreEncryption: true, updateMetadata: false })
  const encryptRef = doc.context.trailerInfo.Encrypt
  decryptStringsOf(
    doc,
    encryptRef instanceof PDFRef ? encryptRef : null,
    pre.topLevel,
    (ref, data) => decryptPayload(derived, ref, data, true),
    entries.encryptMetadata,
  )
  delete (doc.context.trailerInfo as { Encrypt?: unknown }).Encrypt
  return {
    bytes: await doc.save({ useObjectStreams: false }),
    permissions: permissionsFromP(Number(/\/P\s+(-?\d+)/.exec(encrypt.body)?.[1] ?? -1)),
    neededPassword: password.length > 0,
    unlockedAs,
    // the save path re-protects an unchanged document under these entries
    protection: { derived, entries, fileId: rawFileId(text) },
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

  // an empty owner password would hand ownership to anyone; an unset one
  // falls back to the user password (the common single-password file)
  const owner = protection.ownerPassword || protection.userPassword
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

  const entries: EncryptEntries = {
    filter: 'Standard',
    v: 5,
    r: 6,
    length: 256,
    p,
    o,
    u,
    oe,
    ue,
    perms,
    cf: { StdCF: { cfm: 'AESV3', authEvent: 'DocOpen', lengthBytes: 32 } },
    stmF: 'StdCF',
    strF: 'StdCF',
    encryptMetadata: true,
  }
  const encryptRef = context.nextRef()
  context.assign(encryptRef, encryptDictFor(context, entries))
  context.trailerInfo.Encrypt = encryptRef

  if (!context.trailerInfo.ID) {
    const id = randomBytes(16).toString('hex').toUpperCase()
    context.trailerInfo.ID = context.obj([PDFHexString.of(id), PDFHexString.of(id)])
  }

  // both payload classes run through StdCF's AESV3 (StmF/StrF in the entries)
  const derived: R6Encrypt = { revision: 6, fileKey, stmIdentity: false, strIdentity: false }
  transformPayloads(doc, encryptRef, (_ref, data) => encryptPayload(derived, data))
}

/**
 * Protect a document under the source file's OWN protection: the /Encrypt
 * entries the source carried are written back and the same file key encrypts
 * the new bytes, so the user and owner passwords that opened the source still
 * open the output. Only a protection change (a dialog save) writes new entries.
 * Save with useObjectStreams: false, after running pdf-lib's flush.
 */
export function encryptDocumentWithOriginalProtection(
  doc: PDFDocument,
  protection: OriginalProtection,
): void {
  const context = doc.context
  const encryptRef = context.nextRef()
  context.assign(encryptRef, encryptDictFor(context, protection.entries))
  context.trailerInfo.Encrypt = encryptRef
  // A legacy (revision 4 and down) file key is derived from the trailer ID, so
  // the output must keep the source's /ID entries: a fresh ID would make the
  // password the file was sealed with stop working.
  if (protection.fileId.length > 0) {
    context.trailerInfo.ID = context.obj(
      protection.fileId.map((bytes) => PDFHexString.of(hex(bytes))),
    )
  } else if (!context.trailerInfo.ID) {
    const id = randomBytes(16).toString('hex').toUpperCase()
    context.trailerInfo.ID = context.obj([PDFHexString.of(id), PDFHexString.of(id)])
  }
  transformPayloads(
    doc,
    encryptRef,
    (ref, data, isString) => encryptPayloadUnder(protection.derived, ref, data, isString),
    protection.entries.encryptMetadata,
  )
}

/** Encrypt one payload the way its revision decrypts it (inverse of decryptPayload) */
function encryptPayloadUnder(
  derived: Derived,
  ref: PDFRef,
  data: Uint8Array,
  isString: boolean,
): Uint8Array {
  if (derived.revision <= 4) {
    // producers write an empty encrypted payload as empty bytes; AES-CBC would
    // otherwise add an IV plus a padding block for nothing
    if (data.length === 0) return data
    const legacy = derived as LegacyEncrypt
    const cfm = isString ? legacy.stringCfm : legacy.streamCfm
    if (cfm === 'Identity') return data
    if (cfm === 'AESV2') {
      const key = legacyObjectKey(legacy.fileKey, ref, true)
      const iv = randomBytes(16)
      return new Uint8Array([...iv, ...aesCbc('encrypt', 128, key, iv, data, true)])
    }
    return rc4(legacyObjectKey(legacy.fileKey, ref, false), data)
  }
  const r6 = derived as R6Encrypt
  if (isString ? r6.strIdentity : r6.stmIdentity) return data
  return encryptPayload(r6, data)
}

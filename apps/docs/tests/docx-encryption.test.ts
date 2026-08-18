import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  clearPendingDocPassword,
  decryptDocx,
  DocxDecryptError,
  docPasswordFor,
  encryptDocx,
  forgetDocPasswords,
  isDiskEncrypted,
  isEncryptedDocx,
  markDiskEncrypted,
  peekPendingDocPassword,
  rememberDocPassword,
  renameDocPassword,
  setDocPassword,
} from '../src/main/docx-encryption'

const plainDocx = () => readFileSync(join(__dirname, 'pagination-corpus/docx/kitchen-sink.docx'))

describe('isEncryptedDocx', () => {
  it('is false for a plain docx (zip) and random bytes', () => {
    expect(isEncryptedDocx(plainDocx())).toBe(false)
    expect(isEncryptedDocx(Buffer.from('not a docx at all'))).toBe(false)
    expect(isEncryptedDocx(Buffer.alloc(0))).toBe(false)
  })

  it('is true for an ECMA-376 encrypted docx', () => {
    const encrypted = encryptDocx(plainDocx(), 'pw')
    expect(isEncryptedDocx(encrypted)).toBe(true)
  })

  it('is false for a CFB container without an EncryptedPackage stream (legacy .doc shape)', () => {
    const cfbNoPackage = Buffer.concat([
      Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
      Buffer.alloc(512),
    ])
    expect(isEncryptedDocx(cfbNoPackage)).toBe(false)
  })
})

describe('encryptDocx / decryptDocx', () => {
  it('round-trips byte-exact with the right password', async () => {
    const original = plainDocx()
    const encrypted = encryptDocx(original, 'S3cret!密码')
    expect(encrypted.subarray(0, 4)).not.toEqual(original.subarray(0, 4))
    const decrypted = await decryptDocx(encrypted, 'S3cret!密码')
    expect(Buffer.compare(decrypted, original)).toBe(0)
  })

  it('rejects a wrong password with reason wrong-password', async () => {
    const encrypted = encryptDocx(plainDocx(), 'right')
    const err = await decryptDocx(encrypted, 'wrong').catch((e) => e)
    expect(err).toBeInstanceOf(DocxDecryptError)
    expect((err as DocxDecryptError).reason).toBe('wrong-password')
  })

  it('reports unrecognized containers as unsupported', async () => {
    // CFB magic but no usable encryption streams (e.g. proprietary/account encryption)
    const bogus = Buffer.concat([
      Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]),
      Buffer.alloc(4096),
    ])
    const err = await decryptDocx(bogus, 'any').catch((e) => e)
    expect(err).toBeInstanceOf(DocxDecryptError)
    expect((err as DocxDecryptError).reason).toBe('unsupported')
  })
})

// Interop lock: fixtures produced by real Microsoft Office, vendored from the
// reference implementation's test suite (nolze/msoffcrypto-tool, MIT). The agile
// one is >4096 bytes, covering multi-segment package decryption.
describe('real Office files', () => {
  const fixture = (name: string) => readFileSync(join(__dirname, 'encrypted-fixtures', name))

  it('decrypts a real Office agile-encrypted docx byte-exact (multi-segment)', async () => {
    const encrypted = fixture('office-agile-password.docx')
    expect(isEncryptedDocx(encrypted)).toBe(true)
    const decrypted = await decryptDocx(encrypted, 'Password1234_')
    expect(Buffer.compare(decrypted, fixture('office-agile-plain.docx'))).toBe(0)
  })

  it('decrypts a real Office standard-encrypted docx byte-exact', async () => {
    const encrypted = fixture('office-standard-password.docx')
    expect(isEncryptedDocx(encrypted)).toBe(true)
    const decrypted = await decryptDocx(encrypted, 'Password1234_')
    expect(Buffer.compare(decrypted, fixture('office-standard-plain.docx'))).toBe(0)
  })
})

describe('password store', () => {
  it('remembers per renderer + path, and forget drops the whole renderer', () => {
    rememberDocPassword(1, '/a.docx', 'pw-a')
    rememberDocPassword(1, '/b.docx', 'pw-b')
    rememberDocPassword(2, '/a.docx', 'pw-other')
    expect(docPasswordFor(1, '/a.docx')).toBe('pw-a')
    expect(docPasswordFor(1, '/b.docx')).toBe('pw-b')
    expect(docPasswordFor(2, '/a.docx')).toBe('pw-other')
    expect(docPasswordFor(1, '/c.docx')).toBeNull()
    expect(docPasswordFor(1, null)).toBeNull()
    forgetDocPasswords(1)
    expect(docPasswordFor(1, '/a.docx')).toBeNull()
    expect(docPasswordFor(2, '/a.docx')).toBe('pw-other')
  })

  it('setDocPassword sets and clears a path-bound password', () => {
    setDocPassword(3, '/x.docx', 'pw-x')
    expect(docPasswordFor(3, '/x.docx')).toBe('pw-x')
    setDocPassword(3, '/x.docx', null)
    expect(docPasswordFor(3, '/x.docx')).toBeNull()
  })

  it('parks a password for a pathless document until its first save clears it', () => {
    setDocPassword(4, null, 'pw-pending')
    // peek is read-only: a failed save keeps the password for the retry
    expect(peekPendingDocPassword(4)).toBe('pw-pending')
    expect(peekPendingDocPassword(4)).toBe('pw-pending')
    clearPendingDocPassword(4)
    expect(peekPendingDocPassword(4)).toBeNull()
    // clearing before save drops it
    setDocPassword(4, null, 'pw-2')
    setDocPassword(4, null, null)
    expect(peekPendingDocPassword(4)).toBeNull()
    // teardown drops it too
    setDocPassword(5, null, 'pw-3')
    forgetDocPasswords(5)
    expect(peekPendingDocPassword(5)).toBeNull()
  })

  it('follows a file rename, only for the renderer that owns it', () => {
    rememberDocPassword(6, '/old.docx', 'pw-r')
    rememberDocPassword(7, '/other.docx', 'pw-o')
    markDiskEncrypted(6, '/old.docx', true)
    renameDocPassword(6, '/old.docx', '/new.docx')
    expect(docPasswordFor(6, '/old.docx')).toBeNull()
    expect(docPasswordFor(6, '/new.docx')).toBe('pw-r')
    // the on-disk-encrypted flag follows the rename with the password
    expect(isDiskEncrypted(6, '/old.docx')).toBe(false)
    expect(isDiskEncrypted(6, '/new.docx')).toBe(true)
    // renaming a path with no stored password is a no-op
    renameDocPassword(7, '/old.docx', '/new.docx')
    expect(docPasswordFor(7, '/new.docx')).toBeNull()
    expect(docPasswordFor(7, '/other.docx')).toBe('pw-o')
  })

  it('tracks on-disk encryption per renderer + path, dropped on teardown', () => {
    markDiskEncrypted(8, '/enc.docx', true)
    expect(isDiskEncrypted(8, '/enc.docx')).toBe(true)
    expect(isDiskEncrypted(8, '/plain.docx')).toBe(false)
    expect(isDiskEncrypted(9, '/enc.docx')).toBe(false)
    // a plain save turns it off (password removed via the ribbon, then saved)
    markDiskEncrypted(8, '/enc.docx', false)
    expect(isDiskEncrypted(8, '/enc.docx')).toBe(false)
    markDiskEncrypted(8, '/enc.docx', true)
    forgetDocPasswords(8)
    expect(isDiskEncrypted(8, '/enc.docx')).toBe(false)
  })
})

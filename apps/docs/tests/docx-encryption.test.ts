import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  commitDocPasswordSave,
  currentDocPasswordIntentRevision,
  decryptDocx,
  decryptRecoveryCopy,
  discardDocPasswordIntents,
  DocxDecryptError,
  docPasswordFor,
  encryptDocx,
  forgetDocPasswords,
  isDiskEncrypted,
  isEncryptedDocx,
  markDiskEncrypted,
  prepareRecoveryDocx,
  rememberDocPassword,
  renameDocPassword,
  setDocPassword,
  snapshotDocPassword,
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

  it('keeps the current disk password separate until old -> new saves successfully', () => {
    rememberDocPassword(3, '/x.docx', 'old')
    markDiskEncrypted(3, '/x.docx', true)
    setDocPassword(3, '/x.docx', 'new')

    const save = snapshotDocPassword(3, '/x.docx')
    expect(save.password).toBe('new')
    expect(docPasswordFor(3, '/x.docx')).toBe('old')

    expect(commitDocPasswordSave(3, save, '/x.docx')).toBe(false)
    expect(docPasswordFor(3, '/x.docx')).toBe('new')
    expect(isDiskEncrypted(3, '/x.docx')).toBe(true)
    expect(snapshotDocPassword(3, '/x.docx').password).toBe('new')
  })

  it('keeps the current disk password separate until old -> none saves successfully', () => {
    rememberDocPassword(30, '/x.docx', 'old')
    markDiskEncrypted(30, '/x.docx', true)
    setDocPassword(30, '/x.docx', null)

    const save = snapshotDocPassword(30, '/x.docx')
    expect(save.password).toBeNull()
    expect(docPasswordFor(30, '/x.docx')).toBe('old')

    expect(commitDocPasswordSave(30, save, '/x.docx')).toBe(false)
    expect(docPasswordFor(30, '/x.docx')).toBeNull()
    expect(isDiskEncrypted(30, '/x.docx')).toBe(false)
  })

  it('accepting a plain replacement clears stale disk and desired password state', () => {
    rememberDocPassword(37, '/x.docx', 'old')
    setDocPassword(37, '/x.docx', 'desired')
    rememberDocPassword(37, '/x.docx', null)
    markDiskEncrypted(37, '/x.docx', false)
    expect(docPasswordFor(37, '/x.docx')).toBeNull()
    // Auxiliary opens (for example Compare) do not discard the active
    // document's intent; accepting a replacement does so explicitly.
    expect(snapshotDocPassword(37, '/x.docx').password).toBe('desired')
    discardDocPasswordIntents(37)
    expect(snapshotDocPassword(37, '/x.docx').password).toBeNull()
  })

  it('replacement cleanup preserves intents created after its captured cutoff', () => {
    setDocPassword(38, null, 'old-draft')
    setDocPassword(38, '/old.docx', 'old-path')
    const cutoff = currentDocPasswordIntentRevision()

    setDocPassword(38, null, 'replacement-draft')
    setDocPassword(38, '/new.docx', 'replacement-path')
    discardDocPasswordIntents(38, cutoff)

    expect(snapshotDocPassword(38, null).password).toBe('replacement-draft')
    expect(snapshotDocPassword(38, '/new.docx').password).toBe('replacement-path')
    expect(snapshotDocPassword(38, '/old.docx').password).toBeNull()
  })

  it('a failed save leaves desired state pending, and success does not consume a newer intent', () => {
    rememberDocPassword(31, '/x.docx', 'old')
    markDiskEncrypted(31, '/x.docx', true)
    setDocPassword(31, '/x.docx', 'first')

    const failed = snapshotDocPassword(31, '/x.docx')
    // No commit: an atomic-write failure keeps both the disk baseline and intent.
    expect(docPasswordFor(31, '/x.docx')).toBe('old')
    expect(snapshotDocPassword(31, '/x.docx')).toEqual(failed)

    setDocPassword(31, '/x.docx', 'newer')
    expect(commitDocPasswordSave(31, failed, '/x.docx')).toBe(true)
    expect(docPasswordFor(31, '/x.docx')).toBe('first')
    expect(snapshotDocPassword(31, '/x.docx').password).toBe('newer')
  })

  it('parks a password for a pathless document until its first successful save', () => {
    setDocPassword(4, null, 'pw-pending')
    const firstSave = snapshotDocPassword(4, null)
    expect(firstSave.password).toBe('pw-pending')
    expect(snapshotDocPassword(4, null)).toEqual(firstSave)
    expect(commitDocPasswordSave(4, firstSave, '/first.docx')).toBe(false)
    expect(docPasswordFor(4, '/first.docx')).toBe('pw-pending')

    // Clearing before another pathless save makes that save plain.
    setDocPassword(4, null, 'pw-2')
    setDocPassword(4, null, null)
    expect(snapshotDocPassword(4, null).password).toBeNull()
    // teardown drops it too
    setDocPassword(5, null, 'pw-3')
    forgetDocPasswords(5)
    expect(snapshotDocPassword(5, null).password).toBeNull()
  })

  it('moves a pathless or Save As concurrent intent to the landed path', () => {
    setDocPassword(32, null, 'first')
    const firstSave = snapshotDocPassword(32, null)
    setDocPassword(32, null, 'newer')
    expect(commitDocPasswordSave(32, firstSave, '/new.docx')).toBe(true)
    expect(docPasswordFor(32, '/new.docx')).toBe('first')
    expect(snapshotDocPassword(32, '/new.docx').password).toBe('newer')

    rememberDocPassword(33, '/old.docx', 'disk-old')
    setDocPassword(33, '/old.docx', 'save-as')
    const saveAs = snapshotDocPassword(33, '/old.docx')
    setDocPassword(33, '/old.docx', null)
    expect(commitDocPasswordSave(33, saveAs, '/copy.docx')).toBe(true)
    expect(docPasswordFor(33, '/old.docx')).toBe('disk-old')
    expect(docPasswordFor(33, '/copy.docx')).toBe('save-as')
    expect(snapshotDocPassword(33, '/copy.docx').password).toBeNull()
  })

  it('follows a file rename, only for the renderer that owns it', () => {
    rememberDocPassword(6, '/old.docx', 'pw-r')
    setDocPassword(6, '/old.docx', 'pw-next')
    rememberDocPassword(7, '/other.docx', 'pw-o')
    markDiskEncrypted(6, '/old.docx', true)
    renameDocPassword(6, '/old.docx', '/new.docx')
    expect(docPasswordFor(6, '/old.docx')).toBeNull()
    expect(docPasswordFor(6, '/new.docx')).toBe('pw-r')
    expect(snapshotDocPassword(6, '/new.docx').password).toBe('pw-next')
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
    setDocPassword(8, '/enc.docx', 'desired')
    setDocPassword(8, null, 'pathless')
    forgetDocPasswords(8)
    expect(isDiskEncrypted(8, '/enc.docx')).toBe(false)
    expect(docPasswordFor(8, '/enc.docx')).toBeNull()
    expect(snapshotDocPassword(8, '/enc.docx').password).toBeNull()
    expect(snapshotDocPassword(8, null).password).toBeNull()
  })

  it('encrypts and decrypts recovery with the disk password after desired new or none', async () => {
    const plain = plainDocx()
    rememberDocPassword(34, '/enc.docx', 'disk-old')
    markDiskEncrypted(34, '/enc.docx', true)

    setDocPassword(34, '/enc.docx', 'desired-new')
    const whileChanging = prepareRecoveryDocx(34, '/enc.docx', plain)
    expect(whileChanging).not.toBeNull()
    await expect(decryptRecoveryCopy(34, '/enc.docx', whileChanging!)).resolves.toEqual(plain)
    await expect(decryptDocx(whileChanging!, 'desired-new')).rejects.toMatchObject({
      reason: 'wrong-password',
    })

    setDocPassword(34, '/enc.docx', null)
    const whileRemoving = prepareRecoveryDocx(34, '/enc.docx', plain)
    expect(whileRemoving).not.toBeNull()
    await expect(decryptDocx(whileRemoving!, 'disk-old')).resolves.toEqual(plain)
  })

  it('never returns plaintext recovery bytes for an encrypted disk file without its password', () => {
    markDiskEncrypted(35, '/enc.docx', true)
    expect(prepareRecoveryDocx(35, '/enc.docx', plainDocx())).toBeNull()

    markDiskEncrypted(36, '/plain.docx', false)
    setDocPassword(36, '/plain.docx', 'desired-next-save')
    const plain = plainDocx()
    expect(prepareRecoveryDocx(36, '/plain.docx', plain)).toBe(plain)
  })
})

/**
 * The standard security handler, both directions. The fixtures are real
 * producer output — LibreOffice's RC4-128 default and pypdf's AES-128 /
 * AES-256 — because the whole point is agreement with independent
 * implementations; pdf.js is the second oracle for the encryptor's output.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { WrongPasswordError, decryptDocument, encryptDocument } from '../src/main/pdf-crypt'
import { applySaveRequest } from '../src/main/save-pdf'
import type { SavePdfRequest } from '../src/shared/ipc'

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(__dirname, 'fixtures', name)))

const PROBE_TEXT = 'Encrypted Edit Probe'

/** pdf.js is the independent reader; pass a copy, it takes ownership of the buffer */
async function textOf(bytes: Uint8Array, password?: string): Promise<string> {
  const doc = await getDocument({ data: bytes.slice(), password }).promise
  const page = await doc.getPage(1)
  const content = await page.getTextContent()
  return content.items.map((item) => item.str).join(' ')
}

const ENCRYPTED_FIXTURES = ['enc-r3-rc4.pdf', 'enc-r4-aes128.pdf', 'enc-r5-aes256.pdf']

describe('decryptDocument', () => {
  for (const name of ENCRYPTED_FIXTURES) {
    it(`opens a ${name} fixture into working plain bytes`, async () => {
      const { bytes, permissions, neededPassword } = await decryptDocument(fixture(name), 's3cret')
      expect(neededPassword).toBe(true)
      const plain = await PDFDocument.load(bytes, { updateMetadata: false })
      expect(plain.getPageCount()).toBe(1)
      // pdf.js opens it with no password and the words survived the decryption
      const text = await textOf(bytes)
      expect(text).toContain(PROBE_TEXT)
      expect(text).toContain('Second paragraph for text comparison.')
      expect(permissions.printing).toBe(true)
    })

    it(`rejects a wrong password for ${name}`, async () => {
      await expect(decryptDocument(fixture(name), 'wrong')).rejects.toBeInstanceOf(
        WrongPasswordError,
      )
    })
  }

  it('a permissions-only file opens with the empty password', async () => {
    const { bytes, neededPassword, permissions } = await decryptDocument(
      fixture('enc-r5-owneronly.pdf'),
      '',
    )
    expect(neededPassword).toBe(false)
    expect(await textOf(bytes)).toContain(PROBE_TEXT)
    expect(permissions.printing).toBe(true)
  })
})

describe('encryptDocument', () => {
  it('produces AES-256 a reader accepts, with the permissions it was given', async () => {
    const { bytes: plain } = await decryptDocument(fixture('enc-r3-rc4.pdf'), 's3cret')
    const doc = await PDFDocument.load(plain, { updateMetadata: false })
    encryptDocument(doc, {
      userPassword: 'newpw',
      permissions: { printing: true, modifying: false, annotating: true, copying: false },
    })
    const sealed = await doc.save({ useObjectStreams: false })

    // no password, no entry; the right one opens and reads
    await expect(getDocument({ data: sealed.slice() }).promise).rejects.toThrow()
    expect(await textOf(sealed, 'newpw')).toContain(PROBE_TEXT)

    // the protection round-trips through our own opener with its /P intact
    const reopened = await decryptDocument(sealed, 'newpw')
    expect(reopened.permissions.printing).toBe(true)
    expect(reopened.permissions.modifying).toBe(false)
    expect(reopened.permissions.annotating).toBe(true)
    expect(reopened.permissions.copying).toBe(false)
    expect(await textOf(reopened.bytes)).toContain(PROBE_TEXT)
  })

  it('accepts the empty user password (protection an owner can lift)', async () => {
    const doc = await PDFDocument.load(fixture('plain-sample.pdf'), { updateMetadata: false })
    encryptDocument(doc, {
      userPassword: '',
      ownerPassword: 'owner3cret',
      permissions: { printing: true, modifying: true, annotating: true, copying: false },
    })
    const sealed = await doc.save({ useObjectStreams: false })
    expect(await textOf(sealed)).toContain(PROBE_TEXT)
    const reopened = await decryptDocument(sealed, '')
    expect(reopened.neededPassword).toBe(false)
    expect(reopened.permissions.copying).toBe(false)
  })
})

describe('protection through the save pipeline', () => {
  it('a protected save decrypts the source, edits, and re-protects the output', async () => {
    const request: SavePdfRequest = {
      path: 'fixture.pdf',
      markups: [],
      drawings: [],
      formValues: [],
      stamps: [],
      encryption: {
        userPassword: 's3cret',
        ownerPassword: 's3cret',
        permissions: { printing: true, modifying: true, annotating: true, copying: true },
      },
    }
    const { bytes } = await applySaveRequest(fixture('enc-r3-rc4.pdf'), request)
    // still protected, still opening with the same password, edits pipeline unharmed
    await expect(getDocument({ data: bytes.slice() }).promise).rejects.toThrow()
    expect(await textOf(bytes, 's3cret')).toContain(PROBE_TEXT)
    const reopened = await decryptDocument(bytes, 's3cret')
    expect(await textOf(reopened.bytes)).toContain(PROBE_TEXT)
  })

  it('a second protected save opens the previous protected output', async () => {
    const request: SavePdfRequest = {
      path: 'fixture.pdf',
      markups: [],
      drawings: [],
      formValues: [],
      stamps: [],
      encryption: {
        userPassword: 'pw-one',
        ownerPassword: 'pw-one',
        permissions: { printing: true, modifying: true, annotating: true, copying: true },
      },
    }
    const first = (await applySaveRequest(fixture('plain-sample.pdf'), request)).bytes
    const second = await applySaveRequest(first, {
      ...request,
      encryption: {
        userPassword: 'pw-two',
        ownerPassword: 'pw-two',
        permissions: { printing: true, modifying: true, annotating: true, copying: true },
      },
      decryption: { userPassword: 'pw-one' },
    })
    // the disk password changed underneath: the new one opens, the old one does not
    expect(await textOf(second.bytes, 'pw-two')).toContain(PROBE_TEXT)
    await expect(decryptDocument(second.bytes, 'pw-one')).rejects.toBeInstanceOf(WrongPasswordError)
  })

  it('saving with decryption only writes plain bytes', async () => {
    const { bytes } = await applySaveRequest(fixture('enc-r5-aes256.pdf'), {
      path: 'fixture.pdf',
      markups: [],
      drawings: [],
      formValues: [],
      stamps: [],
      decryption: { userPassword: 's3cret' },
    })
    expect(await textOf(bytes)).toContain(PROBE_TEXT)
  })
})

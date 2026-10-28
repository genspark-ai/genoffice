/**
 * The standard security handler's reading half. The fixtures are real producer
 * output — LibreOffice's RC4-128 default and pypdf's AES-128 / AES-256 — because
 * the whole point is agreement with independent implementations; pdf.js is the
 * second oracle for what the decrypted bytes must show.
 */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { WrongPasswordError, decryptDocument } from '../src/main/pdf-crypt'
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

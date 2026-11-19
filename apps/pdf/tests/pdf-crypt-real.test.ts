import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { readPdfText } from '../src/main/read-text'
import {
  CertificateEncryptionError,
  decryptDocument,
  WrongPasswordError,
} from '../src/main/pdf-crypt'

/**
 * Real files: qpdf writes each variant from one base document with object
 * streams on, which is the shape an encrypted PDF actually arrives in —
 * pdf-lib cannot parse one before the streams are decrypted.
 */
const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(__dirname, 'fixtures', name)))

describe('opening a password-protected document', () => {
  it.each(['enc256.pdf', 'enc128.pdf', 'encrc4.pdf'])(
    'decrypts %s and the pages survive',
    async (name) => {
      const out = await decryptDocument(fixture(name), 'user123')
      const doc = await PDFDocument.load(out.bytes)
      expect(doc.getPageCount()).toBe(2)
      // the strongest check there is: a real parser reads the decrypted text
      const text = await readPdfText(new Uint8Array(out.bytes))
      expect(text.pages[0]!.text).toContain('confidential')
    },
  )

  it('rejects a wrong password', async () => {
    await expect(decryptDocument(fixture('enc256.pdf'), 'nope')).rejects.toBeInstanceOf(
      WrongPasswordError,
    )
  })

  it('names certificate encryption instead of asking for a password', async () => {
    // the shell suite's certEncrypted.pdf is Adobe public-key encrypted: no
    // password exists, so "wrong password" would be a lie the user cannot act on
    const bytes = new Uint8Array(
      readFileSync(join(__dirname, '../../shell/tests/fixtures/certEncrypted.pdf')),
    )
    await expect(decryptDocument(bytes, '')).rejects.toBeInstanceOf(CertificateEncryptionError)
  })
})

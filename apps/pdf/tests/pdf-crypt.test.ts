/**
 * The standard security handler, both directions. The fixtures are real
 * producer output — LibreOffice's RC4-128 default and pypdf's AES-128 /
 * AES-256 — because the whole point is agreement with independent
 * implementations; pdf.js is the second oracle for the encryptor's output.
 */
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { PDFDict, PDFDocument, PDFHexString, PDFName, PDFStream, PDFString } from 'pdf-lib'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { WrongPasswordError, decryptDocument, encryptDocument } from '../src/main/pdf-crypt'
import { applySaveRequest, savePdfToPath } from '../src/main/save-pdf'
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
      const { bytes, permissions, neededPassword, unlockedAs } = await decryptDocument(
        fixture(name),
        's3cret',
      )
      expect(neededPassword).toBe(true)
      expect(unlockedAs).toBe('user')
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
    const { bytes, neededPassword, permissions, unlockedAs } = await decryptDocument(
      fixture('enc-r5-owneronly.pdf'),
      '',
    )
    expect(neededPassword).toBe(false)
    // the empty password is the user's: the /P limits bind the session
    expect(unlockedAs).toBe('user')
    expect(await textOf(bytes)).toContain(PROBE_TEXT)
    expect(permissions.printing).toBe(true)
  })

  it('reads the crypt filter /StmF and /StrF name, not the first /CFM in the dict', async () => {
    // a producer that carries an /Identity filter next to StdCF: the streams
    // still run through StdCF's AESV2, whatever appears first in the dict
    const raw = Buffer.from(fixture('enc-r4-aes128.pdf')).toString('latin1')
    const reordered = raw.replace('/CF <<', '/CF << /Crypt << /CFM /Identity >> ')
    expect(reordered).not.toBe(raw)
    const { bytes } = await decryptDocument(
      new Uint8Array(Buffer.from(reordered, 'latin1')),
      's3cret',
    )
    expect(await textOf(bytes)).toContain(PROBE_TEXT)
  })

  it('a V4 file without a root /Length takes the key size from the crypt filter', async () => {
    // /Length is optional; writers that omit it rely on StdCF's /Length 16 (bytes)
    const raw = Buffer.from(fixture('enc-r4-aes128.pdf')).toString('latin1')
    const withoutRootLength = raw.replace('/Length 128\n', '')
    expect(withoutRootLength).not.toBe(raw)
    const source = new Uint8Array(Buffer.from(withoutRootLength, 'latin1'))
    const { bytes } = await decryptDocument(source, 's3cret')
    expect(await textOf(bytes)).toContain(PROBE_TEXT)

    // and an unchanged re-save does not invent a bogus /Length 0
    const saved = await applySaveRequest(
      source,
      saveRequest({
        decryption: { userPassword: 's3cret' },
        protectionDirty: false,
        encryption: {
          userPassword: 's3cret',
          ownerPassword: 's3cret',
          permissions: ALL_PERMISSIONS,
        },
      }),
    )
    expect(Buffer.from(saved.bytes).toString('latin1')).not.toMatch(/\/Length 0\b/)
    expect(await textOf(saved.bytes, 's3cret')).toContain(PROBE_TEXT)
  })

  it('accepts literal-string /O and /U entries', async () => {
    // some producers write the password entries as (…) literals, not <…> hex
    const raw = Buffer.from(fixture('enc-r3-rc4.pdf')).toString('latin1')
    const asLiteral = raw.replace(
      /\/([OU])\s*<([0-9a-fA-F\s]+)>/g,
      (_all, name: string, hex: string) => {
        const bytes = Buffer.from(hex.replace(/\s+/g, ''), 'hex')
        let literal = ''
        for (const byte of bytes) {
          const ch = String.fromCharCode(byte)
          literal += ch === '(' || ch === ')' || ch === '\\' ? `\\${ch}` : ch
        }
        return `/${name} (${literal})`
      },
    )
    expect(asLiteral).not.toBe(raw)
    const { bytes } = await decryptDocument(
      new Uint8Array(Buffer.from(asLiteral, 'latin1')),
      's3cret',
    )
    expect(await textOf(bytes)).toContain(PROBE_TEXT)
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

  it('a protected save encrypts the streams a lazy embed adds', async () => {
    // a stamp's PNG embeds only inside save()/flush(): flush runs before the
    // in-place encryption pass, so those streams must not stay plaintext —
    // the next save re-decrypts this output and would die on them
    const png1x1 =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
    const request: SavePdfRequest = {
      path: 'fixture.pdf',
      markups: [],
      drawings: [],
      formValues: [],
      textInserts: [
        {
          pageIndex: 0,
          origin: [50, 150],
          text: 'stamped by the suite',
          fontSize: 14,
          color: [0.1, 0.2, 0.3],
        },
      ],
      stamps: [
        {
          pageIndex: 0,
          rect: [40, 40, 80, 80],
          image: png1x1,
        },
      ],
      encryption: {
        userPassword: 's3cret',
        ownerPassword: 's3cret',
        permissions: { printing: true, modifying: true, annotating: true, copying: true },
      },
    }
    const { bytes } = await applySaveRequest(fixture('plain-sample.pdf'), request)
    expect(await textOf(bytes, 's3cret')).toContain('stamped by the suite')
    // a plaintext stream object inside it would surface as
    // 'could not decrypt stream object N'
    const again = await applySaveRequest(bytes, { ...request, textInserts: [], stamps: [] })
    expect(await textOf(again.bytes, 's3cret')).toContain('stamped by the suite')
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

/** hex for a PDFHexString built from latin1 text (what the fixtures' entries use) */
const hexOf = (text: string): string => Buffer.from(text, 'latin1').toString('hex')

const ALL_PERMISSIONS = { printing: true, modifying: true, annotating: true, copying: true }
const NO_MODIFY = { printing: true, modifying: false, annotating: true, copying: true }

function saveRequest(overrides: Partial<SavePdfRequest>): SavePdfRequest {
  return {
    path: 'fixture.pdf',
    markups: [],
    drawings: [],
    formValues: [],
    stamps: [],
    ...overrides,
  }
}

describe('payload string coverage', () => {
  it('encrypts strings in stream dictionaries and top-level indirect strings', async () => {
    const doc = await PDFDocument.load(fixture('plain-sample.pdf'), { updateMetadata: false })
    const context = doc.context
    const topLevelText = 'top-level indirect string'
    const topLevelRef = context.nextRef()
    context.assign(topLevelRef, PDFHexString.of(hexOf(topLevelText)))
    // the shape an embedded file takes: /Params /ModDate on the stream's dictionary
    const modDateText = 'D:20240101000000Z'
    const streamRef = context.nextRef()
    context.assign(
      streamRef,
      context.stream(new Uint8Array([1, 2, 3]), {
        Type: 'EmbeddedFile',
        Params: { ModDate: PDFHexString.of(hexOf(modDateText)) },
      }),
    )
    encryptDocument(doc, { userPassword: 'pw', permissions: ALL_PERMISSIONS })
    const sealed = await doc.save({ useObjectStreams: false })
    const raw = Buffer.from(sealed).toString('latin1')
    // neither string is on the wire as plaintext (the higher-level writers could
    // not see either of them before)
    expect(raw).not.toContain(topLevelText)
    expect(raw).not.toContain(modDateText)

    const reopened = await PDFDocument.load((await decryptDocument(sealed, 'pw')).bytes, {
      updateMetadata: false,
    })
    const objects = [...reopened.context.enumerateIndirectObjects()]
    const topLevel = objects.find(
      ([, obj]) =>
        (obj instanceof PDFHexString || obj instanceof PDFString) &&
        Buffer.from(obj.asBytes()).toString('latin1') === topLevelText,
    )
    expect(topLevel).toBeTruthy()
    const stream = objects
      .map(([, obj]) => obj)
      .find(
        (obj) =>
          obj instanceof PDFStream && String(obj.dict.get(PDFName.of('Type'))) === '/EmbeddedFile',
      )
    expect(stream).toBeInstanceOf(PDFStream)
    const params = (stream as PDFStream).dict.lookup(PDFName.of('Params'))
    expect(params).toBeInstanceOf(PDFDict)
    const modDate = (params as PDFDict).lookup(PDFName.of('ModDate'))
    expect(modDate).toBeInstanceOf(PDFHexString)
    expect(Buffer.from((modDate as PDFHexString).asBytes()).toString('latin1')).toBe(modDateText)
  })
})

describe('re-protecting an unchanged protection', () => {
  /** a file sealed with distinct passwords, like a producer's protected export */
  async function sealedWithDistinctPasswords(): Promise<Uint8Array> {
    const doc = await PDFDocument.load(fixture('plain-sample.pdf'), { updateMetadata: false })
    encryptDocument(doc, {
      userPassword: 'user-pw',
      ownerPassword: 'owner-pw',
      permissions: NO_MODIFY,
    })
    return doc.save({ useObjectStreams: false })
  }

  it('keeps the source file user and owner passwords on an unchanged save', async () => {
    const sealed = await sealedWithDistinctPasswords()
    const { bytes } = await applySaveRequest(
      sealed,
      saveRequest({
        decryption: { userPassword: 'user-pw' },
        protectionDirty: false,
        // what the session carries while protection is unchanged: only the one
        // password it was opened with, and its own permission defaults
        encryption: {
          userPassword: 'user-pw',
          ownerPassword: 'user-pw',
          permissions: ALL_PERMISSIONS,
        },
      }),
    )

    // the reviewer's probe: the original owner password still opens the save
    const asOwner = await decryptDocument(bytes, 'owner-pw')
    expect(asOwner.unlockedAs).toBe('owner')
    expect(await textOf(bytes, 'owner-pw')).toContain(PROBE_TEXT)
    // the user password too, with the source /P rather than the session's
    const asUser = await decryptDocument(bytes, 'user-pw')
    expect(asUser.unlockedAs).toBe('user')
    expect(asUser.permissions).toEqual(NO_MODIFY)
    // and the output is still protected for a reader without a password
    await expect(getDocument({ data: bytes.slice() }).promise).rejects.toThrow()
    // the original /Encrypt entries themselves were carried over (new entries
    // would carry fresh salts even for the same passwords)
    const source = await decryptDocument(sealed, 'user-pw')
    const saved = await decryptDocument(bytes, 'user-pw')
    expect(saved.protection?.entries.o).toEqual(source.protection?.entries.o)
    expect(saved.protection?.entries.u).toEqual(source.protection?.entries.u)
    expect(saved.protection?.entries.ue).toEqual(source.protection?.entries.ue)
    expect(saved.protection?.entries.oe).toEqual(source.protection?.entries.oe)
  })

  it('saves an owner-only file without knowing its owner password', async () => {
    // the session holds only the empty user password; preserving the entries is
    // what lets the owner password it never saw survive the save
    const source = fixture('enc-r5-owneronly.pdf')
    const before = await decryptDocument(source, '')
    const { bytes } = await applySaveRequest(
      source,
      saveRequest({
        decryption: { userPassword: '' },
        protectionDirty: false,
        encryption: { userPassword: '', ownerPassword: '', permissions: ALL_PERMISSIONS },
      }),
    )
    const after = await decryptDocument(bytes, '')
    expect(after.permissions).toEqual(before.permissions)
    expect(after.protection?.entries.o).toEqual(before.protection?.entries.o)
    expect(after.protection?.entries.u).toEqual(before.protection?.entries.u)
  })

  it('keeps both passwords when an owner session saves unchanged', async () => {
    const sealed = await sealedWithDistinctPasswords()
    const { bytes } = await applySaveRequest(
      sealed,
      saveRequest({
        // the session that opened it held the owner password
        decryption: { userPassword: 'owner-pw' },
        protectionDirty: false,
        encryption: {
          userPassword: 'owner-pw',
          ownerPassword: 'owner-pw',
          permissions: ALL_PERMISSIONS,
        },
      }),
    )
    // the user password the session never knew still opens the save
    const asUser = await decryptDocument(bytes, 'user-pw')
    expect(asUser.unlockedAs).toBe('user')
    expect(asUser.permissions).toEqual(NO_MODIFY)
    expect((await decryptDocument(bytes, 'owner-pw')).unlockedAs).toBe('owner')
  })

  it.each(ENCRYPTED_FIXTURES)('re-protects %s with its own entries and file key', async (name) => {
    const source = fixture(name)
    const before = await decryptDocument(source, 's3cret')
    const { bytes } = await applySaveRequest(
      source,
      saveRequest({
        decryption: { userPassword: 's3cret' },
        protectionDirty: false,
        encryption: { userPassword: 's3cret', ownerPassword: 'other', permissions: NO_MODIFY },
      }),
    )
    const after = await decryptDocument(bytes, 's3cret')
    expect(after.protection?.entries.o).toEqual(before.protection?.entries.o)
    expect(after.protection?.entries.u).toEqual(before.protection?.entries.u)
    // the source /P survives, so this really is the carried protection and not
    // a freshly derived one (those would carry NO_MODIFY from the session state)
    expect(after.permissions).toEqual(before.permissions)
    expect(await textOf(bytes, 's3cret')).toContain(PROBE_TEXT)
    await expect(decryptDocument(bytes, 'wrong')).rejects.toBeInstanceOf(WrongPasswordError)
  })

  it('writes new entries when the protection is dirty', async () => {
    const sealed = await sealedWithDistinctPasswords()
    const { bytes } = await applySaveRequest(
      sealed,
      saveRequest({
        decryption: { userPassword: 'user-pw' },
        protectionDirty: true,
        encryption: { userPassword: 'new-pw', permissions: ALL_PERMISSIONS },
      }),
    )
    expect(await textOf(bytes, 'new-pw')).toContain(PROBE_TEXT)
    await expect(decryptDocument(bytes, 'user-pw')).rejects.toBeInstanceOf(WrongPasswordError)
  })

  it('verifies content edits against a preserved protection', async () => {
    // the read-back verification opens the SAVED bytes, so it needs the password
    // the preserved output accepts (the disk one), not the session's guesses
    const dir = mkdtempSync(join(tmpdir(), 'gen-pdf-crypt-'))
    const source = join(dir, 'sealed.pdf')
    const target = join(dir, 'saved.pdf')
    writeFileSync(source, await sealedWithDistinctPasswords())
    await savePdfToPath(
      source,
      target,
      saveRequest({
        path: source,
        targetPath: target,
        decryption: { userPassword: 'user-pw' },
        protectionDirty: false,
        encryption: {
          userPassword: 'user-pw',
          ownerPassword: 'user-pw',
          permissions: ALL_PERMISSIONS,
        },
        textInserts: [
          {
            pageIndex: 0,
            origin: [60, 200],
            text: 'verified under preserved protection',
            fontSize: 14,
            color: [0.1, 0.2, 0.3],
          },
        ],
      }),
    )
    const saved = new Uint8Array(readFileSync(target))
    expect(await textOf(saved, 'owner-pw')).toContain('verified under preserved protection')
  })
})

describe('redaction on a protected document', () => {
  it('redacts the plaintext and protects the cleaned copy afterwards', async () => {
    const { bytes } = await applySaveRequest(
      fixture('enc-r5-aes256.pdf'),
      saveRequest({
        decryption: { userPassword: 's3cret' },
        protectionDirty: false,
        encryption: {
          userPassword: 's3cret',
          ownerPassword: 's3cret',
          permissions: ALL_PERMISSIONS,
        },
        redactions: [{ pageIndex: 0, rect: [40, 40, 300, 160] }],
      }),
    )
    // the redaction ran (PDFium can only do that on plaintext) and the output is
    // still protected under the source's own protection, opening with s3cret
    const text = await textOf(bytes, 's3cret')
    expect(text.length).toBeGreaterThan(0)
    expect((await decryptDocument(bytes, 's3cret')).permissions).toBeTruthy()
    await expect(getDocument({ data: bytes.slice() }).promise).rejects.toThrow()
  })

  it('re-protects legacy protection after redaction', async () => {
    const { bytes } = await applySaveRequest(
      fixture('enc-r4-aes128.pdf'),
      saveRequest({
        decryption: { userPassword: 's3cret' },
        protectionDirty: false,
        encryption: { userPassword: 's3cret', ownerPassword: 's3cret', permissions: NO_MODIFY },
        redactions: [{ pageIndex: 0, rect: [40, 40, 300, 160] }],
      }),
    )
    const reopened = await decryptDocument(bytes, 's3cret')
    expect(reopened.permissions).toEqual(
      (await decryptDocument(fixture('enc-r4-aes128.pdf'), 's3cret')).permissions,
    )
  })

  it('writes the new protection after redacting when the protection changed', async () => {
    const { bytes } = await applySaveRequest(
      fixture('enc-r5-aes256.pdf'),
      saveRequest({
        decryption: { userPassword: 's3cret' },
        protectionDirty: true,
        encryption: { userPassword: 'fresh-pw', permissions: ALL_PERMISSIONS },
        redactions: [{ pageIndex: 0, rect: [40, 40, 300, 160] }],
      }),
    )
    expect((await textOf(bytes, 'fresh-pw')).length).toBeGreaterThan(0)
    await expect(decryptDocument(bytes, 's3cret')).rejects.toBeInstanceOf(WrongPasswordError)
  })
})

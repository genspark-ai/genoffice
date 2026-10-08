// @vitest-environment node
import { PDFDocument, degrees } from 'pdf-lib'
import forge from 'node-forge'
import { PNG } from 'pngjs'
import { beforeAll, describe, expect, it } from 'vitest'
import { inspectCertificate, signPdf, SignError } from '../src/main/pdf-sign'
import { readPdfSignatures } from '../src/main/pdf-signatures'

const PASSWORD = 'pw-测试'

let p12: Uint8Array
let otherP12: Uint8Array

function makeP12(commonName: string): Uint8Array {
  const keys = forge.pki.rsa.generateKeyPair(2048)
  const cert = forge.pki.createCertificate()
  cert.publicKey = keys.publicKey
  cert.serialNumber = '01' + forge.util.bytesToHex(forge.random.getBytesSync(8))
  cert.validity.notBefore = new Date(Date.now() - 86_400_000)
  cert.validity.notAfter = new Date(Date.now() + 365 * 86_400_000)
  const attrs = [
    {
      name: 'commonName',
      value: forge.util.encodeUtf8(commonName),
      valueTagClass: forge.asn1.Type.UTF8,
    },
    { name: 'organizationName', value: 'GenOffice' },
  ]
  cert.setSubject(attrs)
  cert.setIssuer(attrs)
  cert.sign(keys.privateKey, forge.md.sha256.create())
  const asn1 = forge.pkcs12.toPkcs12Asn1(keys.privateKey, [cert], PASSWORD, { algorithm: '3des' })
  return Buffer.from(forge.asn1.toDer(asn1).getBytes(), 'binary')
}

async function samplePdf(
  opts: { objectStreams?: boolean; pages?: number; rotate?: number } = {},
): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  for (let i = 0; i < (opts.pages ?? 2); i++) {
    const page = doc.addPage([400, 300])
    page.drawText(`Page ${i + 1}`, { x: 40, y: 200, size: 24 })
    if (opts.rotate) page.setRotation(degrees(opts.rotate))
  }
  return doc.save({ useObjectStreams: opts.objectStreams ?? true })
}

function pngBox(width: number, height: number): Uint8Array {
  const png = new PNG({ width, height })
  for (let i = 0; i < width * height; i++) {
    png.data.writeUInt32BE(0x2563ebff, i * 4)
  }
  return PNG.sync.write(png)
}

beforeAll(() => {
  p12 = makeP12('测试用户 Alice')
  otherP12 = makeP12('Bob')
}, 60_000)

describe('signing and reading back', () => {
  for (const objectStreams of [true, false]) {
    it(`signs a document written ${objectStreams ? 'with an xref stream' : 'with a classic xref table'}`, async () => {
      const signed = await signPdf(await samplePdf({ objectStreams }), {
        p12,
        password: PASSWORD,
        reason: '审批通过',
        location: 'Hong Kong',
      })
      const sigs = await readPdfSignatures(signed)
      expect(sigs).toHaveLength(1)
      const [sig] = sigs
      expect(sig!.fieldName).toBe('Signature1')
      expect(sig!.signer?.commonName).toBe('测试用户 Alice')
      expect(sig!.reason).toBe('审批通过')
      expect(sig!.location).toBe('Hong Kong')
      expect(sig!.coversWholeDocument).toBe(true)
      expect(sig!.signingTimeSigned).toBe(true)
      expect(sig!.digestAlgorithm).toBe('SHA256')
      // intact but nobody vouches for a throwaway self-signed certificate
      expect(sig!.problems).toEqual(['self-signed'])
      expect(sig!.status).toBe('valid')
      // the document is still a loadable PDF with both pages
      expect((await PDFDocument.load(signed)).getPageCount()).toBe(2)
    })
  }

  it("names the signer by the certificate's own issuer encoding, so other verifiers can find it", async () => {
    const signed = Buffer.from(await signPdf(await samplePdf(), { p12, password: PASSWORD }))
    const hex = /\/Contents <([0-9A-F]+)>/.exec(signed.toString('latin1'))![1]!
    const der = Buffer.from(hex, 'hex')
    const root = forge.asn1.fromDer(forge.util.createBuffer(der.toString('binary')), {
      parseAllBytes: false,
    })
    const signedData = ((root.value as forge.asn1.Asn1[])[1]!.value as forge.asn1.Asn1[])[0]!
    const parts = signedData.value as forge.asn1.Asn1[]
    const certNode = (
      parts.find((n) => n.tagClass === forge.asn1.Class.CONTEXT_SPECIFIC)!
        .value as forge.asn1.Asn1[]
    )[0]!
    const tbs = (certNode.value as forge.asn1.Asn1[])[0]!.value as forge.asn1.Asn1[]
    const certIssuer = forge.asn1
      .toDer(tbs[tbs[0]!.tagClass === forge.asn1.Class.CONTEXT_SPECIFIC ? 3 : 2]!)
      .getBytes()
    const signerInfo = (parts[parts.length - 1]!.value as forge.asn1.Asn1[])[0]!
    const idAndSerial = (signerInfo.value as forge.asn1.Asn1[])[1]!
    const signerIssuer = forge.asn1.toDer((idAndSerial.value as forge.asn1.Asn1[])[0]!).getBytes()
    expect(signerIssuer).toBe(certIssuer)
  })

  it('still calls a self-signed certificate without certificate-signing key usage self-signed', async () => {
    // Windows New-SelfSignedCertificate produces these: digitalSignature only, no keyCertSign
    const keys = forge.pki.rsa.generateKeyPair(2048)
    const cert = forge.pki.createCertificate()
    cert.publicKey = keys.publicKey
    cert.serialNumber = '07'
    cert.validity.notBefore = new Date(Date.now() - 86_400_000)
    cert.validity.notAfter = new Date(Date.now() + 86_400_000)
    const attrs = [{ name: 'commonName', value: 'Windows Style' }]
    cert.setSubject(attrs)
    cert.setIssuer(attrs)
    cert.setExtensions([{ name: 'keyUsage', digitalSignature: true, critical: true }])
    cert.sign(keys.privateKey, forge.md.sha256.create())
    const asn1 = forge.pkcs12.toPkcs12Asn1(keys.privateKey, [cert], PASSWORD, { algorithm: '3des' })
    const winP12 = Buffer.from(forge.asn1.toDer(asn1).getBytes(), 'binary')
    const signed = await signPdf(await samplePdf(), { p12: winP12, password: PASSWORD })
    const [sig] = await readPdfSignatures(signed)
    expect(sig!.signer?.selfSigned).toBe(true)
    expect(sig!.problems).toEqual(['self-signed'])
    expect(sig!.status).toBe('valid')
  })

  it('flags a document altered inside the signed bytes', async () => {
    const signed = Buffer.from(await signPdf(await samplePdf(), { p12, password: PASSWORD }))
    const at = signed.indexOf('%PDF-') + 20
    signed[at] = signed[at]! ^ 0xff
    const [sig] = await readPdfSignatures(signed)
    expect(sig!.status).toBe('invalid')
    expect(sig!.problems).toContain('digest-mismatch')
  })

  it('flags a signature whose CMS blob was swapped for another signer', async () => {
    const a = Buffer.from(await signPdf(await samplePdf(), { p12, password: PASSWORD }))
    const b = Buffer.from(
      await signPdf(await samplePdf({ pages: 3 }), { p12: otherP12, password: PASSWORD }),
    )
    const hexOf = (file: Buffer) => /\/Contents <([0-9A-F]+)>/.exec(file.toString('latin1'))![1]!
    const swapped = Buffer.from(
      a
        .toString('latin1')
        .replace(hexOf(a), hexOf(b).padEnd(hexOf(a).length, '0').slice(0, hexOf(a).length)),
      'latin1',
    )
    const [sig] = await readPdfSignatures(swapped)
    expect(sig!.status).toBe('invalid')
  })

  it('keeps an earlier signature intact when a second one is appended', async () => {
    const first = await signPdf(await samplePdf(), { p12, password: PASSWORD })
    const second = await signPdf(first, { p12: otherP12, password: PASSWORD })
    const sigs = await readPdfSignatures(second)
    expect(sigs.map((s) => s.fieldName)).toEqual(['Signature1', 'Signature2'])
    expect(sigs[0]!.problems).toContain('document-changed-after-signing')
    expect(sigs[0]!.problems).not.toContain('digest-mismatch')
    expect(sigs[0]!.status).toBe('warning')
    expect(sigs[1]!.coversWholeDocument).toBe(true)
    expect(sigs[1]!.signer?.commonName).toBe('Bob')
  })

  it('reports content appended after a signature as a change, not as tampering', async () => {
    const signed = await signPdf(await samplePdf({ objectStreams: false }), {
      p12,
      password: PASSWORD,
    })
    const extended = Buffer.concat([Buffer.from(signed), Buffer.from('\n% appended\n')])
    const [sig] = await readPdfSignatures(extended)
    expect(sig!.coversWholeDocument).toBe(false)
    expect(sig!.problems).toContain('document-changed-after-signing')
    expect(sig!.problems).not.toContain('digest-mismatch')
  })

  it('reports no signatures for an unsigned or non-PDF input', async () => {
    expect(await readPdfSignatures(await samplePdf())).toEqual([])
    expect(await readPdfSignatures(Buffer.from('not a pdf /ByteRange'))).toEqual([])
  })
})

describe('certifying signatures', () => {
  it('records the DocMDP level and refuses to certify an already signed document', async () => {
    const certified = await signPdf(await samplePdf(), { p12, password: PASSWORD, certifyLevel: 2 })
    const [sig] = await readPdfSignatures(certified)
    expect(sig!.certifiedLevel).toBe(2)
    const doc = await PDFDocument.load(certified)
    expect(
      doc.catalog.has(forge ? (await import('pdf-lib')).PDFName.of('Perms') : (undefined as never)),
    ).toBe(true)
    await expect(
      signPdf(await signPdf(await samplePdf(), { p12, password: PASSWORD }), {
        p12: otherP12,
        password: PASSWORD,
        certifyLevel: 2,
      }),
    ).rejects.toBeInstanceOf(SignError)
  })
})

describe('visible signatures', () => {
  for (const rotate of [0, 90, 180, 270]) {
    it(`places an appearance on a page rotated ${rotate}°`, async () => {
      const signed = await signPdf(await samplePdf({ rotate, pages: 1 }), {
        p12,
        password: PASSWORD,
        visible: {
          pageIndex: 0,
          position: 'bottom-right',
          width: 160,
          height: 60,
          png: pngBox(32, 12),
        },
      })
      const [sig] = await readPdfSignatures(signed)
      expect(sig!.pageIndex).toBe(0)
      expect(sig!.problems).toEqual(['self-signed'])
    })
  }
})

describe('certificate files', () => {
  it('describes the signer', () => {
    const result = inspectCertificate(p12, PASSWORD)
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.signer.commonName).toBe('测试用户 Alice')
      expect(result.signer.keyAlgorithm).toBe('RSA 2048')
      expect(result.signer.selfSigned).toBe(true)
    }
  })

  it('distinguishes a wrong password from a damaged file', () => {
    expect(inspectCertificate(p12, 'nope')).toEqual({ ok: false, error: 'cert-password' })
    expect(inspectCertificate(Buffer.from('junk'), PASSWORD)).toEqual({
      ok: false,
      error: 'cert-invalid',
    })
  })
})

// @vitest-environment node
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import forge from 'node-forge'
import { PDFDocument } from 'pdf-lib'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import { PDF_CHANNELS } from '../src/shared/ipc'
import type { SignWithCertificateResult } from '../src/shared/ipc'
import { readPdfSignatures } from '../src/main/pdf-signatures'

vi.mock('../src/main/system-certificates', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/main/system-certificates')>()),
  listSystemCertificates: vi.fn(),
  exportSystemCertificate: vi.fn(),
}))

type Handler = (event: { sender: FakeContents }, ...args: any[]) => any
const handlers = new Map<string, Handler>()
const pickOpen = vi.fn()
const pickSave = vi.fn()
let lastContents: FakeContents
let nextId = 100
class FakeContents {
  id = ++nextId
  listeners = new Map<string, () => void>()
  once = (event: string, fn: () => void) => this.listeners.set(event, fn)
  on = this.once
  setWindowOpenHandler = vi.fn()
  loadURL = vi.fn()
  loadFile = vi.fn()
  isDestroyed = () => false
  send = vi.fn()
}
vi.mock('electron', () => ({
  app: { on: vi.fn(), whenReady: () => new Promise(() => {}) },
  dialog: {
    showOpenDialog: (...args: unknown[]) => pickOpen(...args),
    showSaveDialog: (...args: unknown[]) => pickSave(...args),
  },
  shell: {},
  BrowserWindow: { fromWebContents: () => null },
  WebContentsView: class {
    webContents = (lastContents = new FakeContents())
  },
  ipcMain: {
    handle: (channel: string, fn: Handler) => handlers.set(channel, fn),
    on: (channel: string, fn: Handler) => handlers.set(channel, fn),
    removeHandler: vi.fn(),
  },
}))
import { createPdfView, setPdfRedactionSavedHook } from '../src/main/pdf-main'
import {
  SystemStoreError,
  exportSystemCertificate,
  listSystemCertificates,
} from '../src/main/system-certificates'

const PASSWORD = 'pw'
let p12: Uint8Array
const dirs: string[] = []

beforeAll(() => {
  const keys = forge.pki.rsa.generateKeyPair(2048)
  const cert = forge.pki.createCertificate()
  cert.publicKey = keys.publicKey
  cert.serialNumber = '01'
  cert.validity.notBefore = new Date(Date.now() - 86_400_000)
  cert.validity.notAfter = new Date(Date.now() + 86_400_000)
  const attrs = [{ name: 'commonName', value: 'Handler Test' }]
  cert.setSubject(attrs)
  cert.setIssuer(attrs)
  cert.sign(keys.privateKey, forge.md.sha256.create())
  const asn1 = forge.pkcs12.toPkcs12Asn1(keys.privateKey, [cert], PASSWORD)
  p12 = Buffer.from(forge.asn1.toDer(asn1).getBytes(), 'binary')
}, 60_000)

afterEach(async () => {
  pickOpen.mockReset()
  pickSave.mockReset()
  setPdfRedactionSavedHook(() => {})
  await Promise.all(dirs.splice(0).map((dir) => rm(dir, { recursive: true, force: true })))
})

async function setup() {
  const dir = await mkdtemp(join(tmpdir(), 'cert-sign-'))
  dirs.push(dir)
  const source = join(dir, 'doc.pdf')
  const certFile = join(dir, 'me.p12')
  const signedCopy = join(dir, 'doc-signed.pdf')
  const pdf = await PDFDocument.create()
  pdf.addPage([300, 200])
  const original = await pdf.save()
  await writeFile(source, original)
  await writeFile(certFile, p12)
  createPdfView(source)
  const wc = lastContents
  const call = (channel: string, ...args: unknown[]) =>
    handlers.get(channel)!({ sender: wc }, ...args)
  pickOpen.mockResolvedValue({ canceled: false, filePaths: [certFile] })
  const { certId } = await call(PDF_CHANNELS.pickCertificate)
  const sign = (over: Record<string, unknown> = {}): Promise<SignWithCertificateResult> =>
    call(PDF_CHANNELS.signWithCertificate, {
      path: source,
      certId,
      password: PASSWORD,
      reason: 'Approved',
      ...over,
    })
  return { dir, source, certId, signedCopy, original, wc, call, sign }
}

describe('certificate signing handlers', () => {
  it('never reveals the certificate path and checks the password', async () => {
    const s = await setup()
    expect(typeof s.certId).toBe('string')
    expect(s.certId).not.toContain('me.p12')
    const good = await s.call(PDF_CHANNELS.inspectCertificate, s.certId, PASSWORD)
    expect(good).toMatchObject({ ok: true, signer: { commonName: 'Handler Test' } })
    expect(await s.call(PDF_CHANNELS.inspectCertificate, s.certId, 'wrong')).toEqual({
      ok: false,
      error: 'cert-password',
    })
    expect(await s.call(PDF_CHANNELS.inspectCertificate, 'unknown-id', PASSWORD)).toEqual({
      ok: false,
      error: 'cert-invalid',
    })
  })

  it('writes a signed copy, leaves the original untouched and moves the view to the copy', async () => {
    const s = await setup()
    const hook = vi.fn()
    setPdfRedactionSavedHook(hook)
    pickSave.mockResolvedValue({ canceled: false, filePath: s.signedCopy })
    const result = await s.sign()
    expect(result).toEqual({ ok: true, path: s.signedCopy })
    expect(new Uint8Array(await readFile(s.source))).toEqual(s.original)
    const [sig] = await readPdfSignatures(new Uint8Array(await readFile(s.signedCopy)))
    expect(sig).toMatchObject({ reason: 'Approved', coversWholeDocument: true })
    expect(sig!.problems).toEqual(['self-signed'])
    expect(s.call(PDF_CHANNELS.consumePending)).toBe(s.signedCopy)
    expect(hook).toHaveBeenCalledWith(s.wc, s.signedCopy)
    await expect(s.call(PDF_CHANNELS.readFile, s.source)).rejects.toThrow('not granted')
    expect(await s.call(PDF_CHANNELS.listDigitalSignatures, s.signedCopy)).toHaveLength(1)
  })

  it('reports a cancelled save dialog without touching anything', async () => {
    const s = await setup()
    pickSave.mockResolvedValue({ canceled: true })
    expect(await s.sign()).toEqual({ ok: false, cancelled: true })
    expect(s.call(PDF_CHANNELS.consumePending)).toBe(s.source)
  })

  it('refuses to overwrite the source file', async () => {
    const s = await setup()
    pickSave.mockResolvedValue({ canceled: false, filePath: s.source })
    expect(await s.sign()).toEqual({ ok: false, cancelled: true })
    expect(new Uint8Array(await readFile(s.source))).toEqual(s.original)
  })

  it('fails on a wrong password before asking where to save', async () => {
    const s = await setup()
    expect(await s.sign({ password: 'wrong' })).toEqual({ ok: false, error: 'cert-password' })
    expect(pickSave).not.toHaveBeenCalled()
  })

  it('refuses while edits are unsaved, for ungranted paths and for malformed requests', async () => {
    const s = await setup()
    handlers.get(PDF_CHANNELS.dirtyChanged)!({ sender: s.wc }, true)
    expect(await s.sign()).toMatchObject({ ok: false, error: 'sign-failed' })
    handlers.get(PDF_CHANNELS.dirtyChanged)!({ sender: s.wc }, false)
    expect(await s.sign({ path: '/etc/passwd' })).toMatchObject({ ok: false, error: 'sign-failed' })
    for (const visible of [
      { pageIndex: -1, position: 'center', width: 100, height: 40, png: 'AA==' },
      { pageIndex: 0, position: 'nowhere', width: 100, height: 40, png: 'AA==' },
      { pageIndex: 0, position: 'center', width: 5, height: 40, png: 'AA==' },
      { pageIndex: 0, position: 'center', width: 100, height: 40, png: 5 },
    ]) {
      expect(await s.sign({ visible })).toMatchObject({ ok: false, error: 'sign-failed' })
    }
    expect(await s.sign({ certifyLevel: 9 })).toMatchObject({ ok: false, error: 'sign-failed' })
    expect(pickSave).not.toHaveBeenCalled()
  })

  describe('system store identities', () => {
    const ref = { source: 'nss' as const, dir: '/home/u/.pki/nssdb', nickname: 'work' }
    const listed = {
      ref,
      label: 'work',
      commonName: 'Handler Test',
      issuerCommonName: 'Handler Test',
      validTo: '2099-01-01T00:00:00.000Z',
      expired: false,
    }

    it('hands the renderer opaque ids, signs with the exported key and drops it afterwards', async () => {
      const s = await setup()
      vi.mocked(listSystemCertificates).mockResolvedValue({ certs: [listed], issues: [] })
      vi.mocked(exportSystemCertificate).mockImplementation(async () => ({
        p12: new Uint8Array(p12),
        password: PASSWORD,
      }))
      const listing = await s.call(PDF_CHANNELS.listSystemCertificates)
      expect(listing.certs).toHaveLength(1)
      const [item] = listing.certs
      expect(JSON.stringify(item)).not.toContain('nssdb')

      expect(await s.call(PDF_CHANNELS.inspectCertificate, item.certId, '')).toMatchObject({
        ok: true,
        signer: { commonName: 'Handler Test' },
      })
      expect(exportSystemCertificate).toHaveBeenCalledTimes(1)
      pickSave.mockResolvedValue({ canceled: false, filePath: s.signedCopy })
      expect(await s.sign({ certId: item.certId, password: '' })).toEqual({
        ok: true,
        path: s.signedCopy,
      })
      // The inspect-time export was reused for the signature, then forgotten
      expect(exportSystemCertificate).toHaveBeenCalledTimes(1)
      const [sig] = await readPdfSignatures(new Uint8Array(await readFile(s.signedCopy)))
      expect(sig!.signer?.commonName).toBe('Handler Test')
    })

    it('reports why a store identity cannot be used', async () => {
      const s = await setup()
      vi.mocked(listSystemCertificates).mockResolvedValue({ certs: [listed], issues: [] })
      const { certs } = await s.call(PDF_CHANNELS.listSystemCertificates)
      vi.mocked(exportSystemCertificate).mockRejectedValue(
        new SystemStoreError('cert-not-exportable'),
      )
      expect(await s.call(PDF_CHANNELS.inspectCertificate, certs[0].certId, '')).toEqual({
        ok: false,
        error: 'cert-not-exportable',
      })
      expect(await s.sign({ certId: certs[0].certId, password: '' })).toEqual({
        ok: false,
        error: 'cert-not-exportable',
      })
      expect(pickSave).not.toHaveBeenCalled()
    })

    it('forgets ids from an earlier listing when the store is listed again', async () => {
      const s = await setup()
      vi.mocked(listSystemCertificates).mockResolvedValue({ certs: [listed], issues: [] })
      const first = await s.call(PDF_CHANNELS.listSystemCertificates)
      await s.call(PDF_CHANNELS.listSystemCertificates)
      expect(await s.call(PDF_CHANNELS.inspectCertificate, first.certs[0].certId, '')).toEqual({
        ok: false,
        error: 'cert-invalid',
      })
    })
  })
})

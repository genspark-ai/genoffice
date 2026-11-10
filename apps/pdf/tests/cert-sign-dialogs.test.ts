/**
 * Certificate signing dialog flow and the signature panel, driven through a fake
 * window.pdfApi (the real one is the main-process handlers covered by pdf-signatures.test.ts).
 */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest'
import type { CertificateInfo, PdfSignatureInfo } from '../src/shared/ipc'
import type { TFunc } from '../src/renderer/i18n/locale'
import { CertSignDialog } from '../src/renderer/CertSignDialog'
import { SignaturesDialog, overallSignatureStatus } from '../src/renderer/SignaturesDialog'

vi.mock('../src/renderer/cert-appearance', () => ({
  renderSignatureAppearance: vi.fn(() => 'UE5H'),
}))

// Identity translator with params appended, so assertions can see what was interpolated
const t = ((key: string, params?: Record<string, string>) =>
  params ? `${key} ${JSON.stringify(params)}` : key) as unknown as TFunc

const cert: CertificateInfo = {
  commonName: 'Alice',
  subject: 'CN=Alice',
  issuerCommonName: 'Alice',
  issuer: 'CN=Alice',
  serialNumber: '01',
  validFrom: '2026-01-01T00:00:00.000Z',
  validTo: '2027-01-01T00:00:00.000Z',
  fingerprint: 'AA:BB',
  keyAlgorithm: 'RSA 2048',
  selfSigned: true,
}

const signature = (over: Partial<PdfSignatureInfo> = {}): PdfSignatureInfo => ({
  fieldName: 'Signature1',
  pageIndex: 1,
  kind: 'signature',
  status: 'valid',
  problems: [],
  subFilter: 'adbe.pkcs7.detached',
  signingTimeSigned: true,
  coversWholeDocument: true,
  signer: cert,
  chain: [cert],
  trusted: false,
  ...over,
})

let root: Root | null = null
let container: HTMLDivElement | null = null

beforeAll(() => {
  ;(globalThis as Record<string, unknown>).IS_REACT_ACT_ENVIRONMENT = true
})

afterEach(async () => {
  if (root) await act(async () => root?.unmount())
  container?.remove()
  root = null
  container = null
  delete (window as unknown as Record<string, unknown>).pdfApi
})

async function mount(element: ReturnType<typeof createElement>) {
  container = document.createElement('div')
  document.body.appendChild(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(element)
    await Promise.resolve()
  })
  return container
}

const flush = () => act(async () => void (await Promise.resolve()))

function setValue(input: HTMLInputElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!
  setter.call(input, value)
  input.dispatchEvent(new Event('input', { bubbles: true }))
}

const buttonWithText = (root: HTMLElement, text: string) =>
  [...root.querySelectorAll<HTMLButtonElement>('button')].find((b) => b.textContent === text)!

function installApi(api: Record<string, unknown>) {
  ;(window as unknown as Record<string, unknown>).pdfApi = {
    listSystemCertificates: async () => ({ certs: [], issues: [] }),
    releaseCertificate: () => {},
    ...api,
  }
}

/** The dialog opens on the system store; the file-based flows start by switching source */
async function useFileSource(dialog: HTMLElement) {
  const select = dialog.querySelector<HTMLSelectElement>('select')!
  await act(async () => {
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value')!.set!.call(select, 'file')
    select.dispatchEvent(new Event('change', { bubbles: true }))
  })
  await flush()
}

describe('CertSignDialog', () => {
  const props = {
    t,
    filePath: '/docs/a.pdf',
    pageCount: 3,
    currentPage: 2,
    hasSignatures: false,
    onCancel: () => {},
  }

  it('cannot sign until the certificate has been checked, then signs with the entered details', async () => {
    const signWithCertificate = vi.fn(async () => ({
      ok: true as const,
      path: '/docs/a-signed.pdf',
    }))
    installApi({
      pickCertificate: async () => ({ certId: 'c1', fileName: 'me.p12' }),
      inspectCertificate: async () => ({ ok: true, signer: cert, chain: [cert] }),
      signWithCertificate,
    })
    const onSigned = vi.fn()
    const dialog = await mount(createElement(CertSignDialog, { ...props, onSigned }))
    await useFileSource(dialog)
    const signButton = buttonWithText(dialog, 'certSignButton')
    expect(signButton.disabled).toBe(true)

    await act(async () => buttonWithText(dialog, 'certChoose').click())
    await flush()
    const password = dialog.querySelector<HTMLInputElement>('input[type="password"]')!
    expect(password.disabled).toBe(false)
    await act(async () => setValue(password, 'secret'))
    await act(async () => buttonWithText(dialog, 'certCheck').click())
    await flush()
    expect(dialog.textContent).toContain('Alice')
    expect(buttonWithText(dialog, 'certSignButton').disabled).toBe(false)

    const [reason] = dialog.querySelectorAll<HTMLInputElement>('input.pdf-modal-input:not([type])')
    await act(async () => setValue(reason!, '  Approved  '))
    await act(async () => buttonWithText(dialog, 'certSignButton').click())
    await flush()

    expect(signWithCertificate).toHaveBeenCalledTimes(1)
    expect(signWithCertificate).toHaveBeenCalledWith({
      path: '/docs/a.pdf',
      certId: 'c1',
      password: 'secret',
      reason: 'Approved',
      location: undefined,
      contactInfo: undefined,
      certifyLevel: undefined,
      visible: { pageIndex: 1, position: 'bottom-right', width: 210, height: 70, png: 'UE5H' },
    })
    expect(onSigned).toHaveBeenCalledWith('/docs/a-signed.pdf')
  })

  it('shows the localized reason when the password is wrong and keeps signing disabled', async () => {
    installApi({
      pickCertificate: async () => ({ certId: 'c1', fileName: 'me.p12' }),
      inspectCertificate: async () => ({ ok: false, error: 'cert-password' }),
    })
    const dialog = await mount(createElement(CertSignDialog, { ...props, onSigned: () => {} }))
    await useFileSource(dialog)
    await act(async () => buttonWithText(dialog, 'certChoose').click())
    await flush()
    await act(async () =>
      setValue(dialog.querySelector<HTMLInputElement>('input[type="password"]')!, 'nope'),
    )
    await act(async () => buttonWithText(dialog, 'certCheck').click())
    await flush()
    expect(dialog.querySelector('[role="alert"]')?.textContent).toBe('certErrPassword')
    expect(buttonWithText(dialog, 'certSignButton').disabled).toBe(true)
  })

  it('does not close or report an error when the save dialog is cancelled', async () => {
    installApi({
      pickCertificate: async () => ({ certId: 'c1', fileName: 'me.p12' }),
      inspectCertificate: async () => ({ ok: true, signer: cert, chain: [cert] }),
      signWithCertificate: async () => ({ ok: false, cancelled: true }),
    })
    const onSigned = vi.fn()
    const dialog = await mount(createElement(CertSignDialog, { ...props, onSigned }))
    await useFileSource(dialog)
    await act(async () => buttonWithText(dialog, 'certChoose').click())
    await flush()
    await act(async () => buttonWithText(dialog, 'certCheck').click())
    await flush()
    await act(async () => buttonWithText(dialog, 'certSignButton').click())
    await flush()
    expect(onSigned).not.toHaveBeenCalled()
    expect(dialog.querySelector('[role="alert"]')).toBeNull()
  })

  it('signs with an identity from the system store without asking for a password', async () => {
    const signWithCertificate = vi.fn(async () => ({
      ok: true as const,
      path: '/docs/a-signed.pdf',
    }))
    const inspectCertificate = vi.fn(async () => ({
      ok: true as const,
      signer: cert,
      chain: [cert],
    }))
    const releaseCertificate = vi.fn()
    installApi({
      listSystemCertificates: async () => ({
        certs: [
          {
            certId: 's1',
            label: 'Alice',
            commonName: 'Alice',
            issuerCommonName: 'Alice',
            validTo: '2027-01-01T00:00:00.000Z',
            expired: false,
          },
        ],
        issues: [],
      }),
      inspectCertificate,
      signWithCertificate,
      releaseCertificate,
    })
    const onSigned = vi.fn()
    const dialog = await mount(createElement(CertSignDialog, { ...props, onSigned }))
    await flush()
    expect(dialog.querySelector('input[type="password"]')).toBeNull()
    await act(async () => buttonWithText(dialog, 'certCheck').click())
    await flush()
    expect(inspectCertificate).toHaveBeenCalledWith('s1', '', undefined)
    await act(async () => buttonWithText(dialog, 'certSignButton').click())
    await flush()
    expect(signWithCertificate).toHaveBeenCalledWith(
      expect.objectContaining({ certId: 's1', password: '' }),
    )
    expect(onSigned).toHaveBeenCalledWith('/docs/a-signed.pdf')
    await act(async () => root!.unmount())
    root = null
    expect(releaseCertificate).toHaveBeenCalledWith('s1')
  })

  it('explains an empty store and a missing NSS tool', async () => {
    installApi({
      listSystemCertificates: async () => ({ certs: [], issues: ['nss-tools-missing'] }),
    })
    const dialog = await mount(createElement(CertSignDialog, { ...props, onSigned: () => {} }))
    await flush()
    expect(dialog.textContent).toContain('certStoreNoTools')
    expect(buttonWithText(dialog, 'certCheck').disabled).toBe(true)
  })

  it('asks for the store password when the store is locked and retries with it', async () => {
    const listSystemCertificates = vi.fn(async () => ({
      certs: [],
      issues: ['store-locked' as const],
    }))
    installApi({ listSystemCertificates })
    const dialog = await mount(createElement(CertSignDialog, { ...props, onSigned: () => {} }))
    await flush()
    await act(async () =>
      setValue(dialog.querySelector<HTMLInputElement>('input[type="password"]')!, 'pw'),
    )
    await act(async () => buttonWithText(dialog, 'certStoreUnlock').click())
    await flush()
    expect(listSystemCertificates).toHaveBeenLastCalledWith('pw')
  })

  it('offers certification only for a document that has no signatures yet', async () => {
    installApi({})
    const dialog = await mount(
      createElement(CertSignDialog, { ...props, hasSignatures: true, onSigned: () => {} }),
    )
    const checkboxes = dialog.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')
    expect(checkboxes[1]!.disabled).toBe(true)
  })
})

describe('SignaturesDialog', () => {
  it('summarizes the worst status across signatures', () => {
    expect(overallSignatureStatus([signature(), signature({ status: 'warning' })])).toBe('warning')
    expect(
      overallSignatureStatus([signature({ status: 'warning' }), signature({ status: 'invalid' })]),
    ).toBe('invalid')
    expect(overallSignatureStatus([signature()])).toBe('valid')
  })

  it('lists each signature with its problems and jumps to its page', async () => {
    const onGoToPage = vi.fn()
    const onClose = vi.fn()
    const dialog = await mount(
      createElement(SignaturesDialog, {
        t,
        onClose,
        onGoToPage,
        signatures: [
          signature(),
          signature({
            fieldName: 'Signature2',
            status: 'invalid',
            problems: ['digest-mismatch'],
            pageIndex: -1,
          }),
        ],
      }),
    )
    const cards = dialog.querySelectorAll('.pdf-sig-card')
    expect(cards).toHaveLength(2)
    expect(cards[0]!.classList.contains('valid')).toBe(true)
    expect(cards[1]!.classList.contains('invalid')).toBe(true)
    expect(cards[1]!.textContent).toContain('sigProblemDigestMismatch')
    // A signature without a page widget has nothing to jump to
    expect(cards[1]!.querySelector('button')).toBeNull()
    await act(async () => cards[0]!.querySelector('button')!.click())
    expect(onGoToPage).toHaveBeenCalledWith(1)
    expect(onClose).toHaveBeenCalled()
  })
})

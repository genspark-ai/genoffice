import { fileURLToPath } from 'node:url'
import { describe, expect, it, vi } from 'vitest'
// same relative path as src/main/pdf2docx-local.ts so PdfLoadError instanceof
// checks see one module instance
import { PdfLoadError } from '../../../packages/pdf2docx/src'
import {
  convertPdfFileToDocxLocal,
  convertPdfFileToDocxLocalWithPrompt,
  convertWithPasswordRetry,
} from '../src/main/pdf2docx-local'

/**
 * The password retry state machine behind the P23 prompt dialog
 * (src/main/pdf2docx-local.ts): password-required loops through the prompt
 * until success or cancel; everything else passes through untouched. The
 * integration tests drive the real converter (PDFium wasm) against
 * tests/fixtures/testPassword4Spaces.pdf, whose user password is four spaces
 * (Apache Tika test corpus).
 */

const passwordRequired = () => new PdfLoadError('password-required', 4)
const FIXTURE = fileURLToPath(new URL('./fixtures/testPassword4Spaces.pdf', import.meta.url))
/** certificate-encrypted (PDFBox AES128ExposedMetadata) → FPDF error 5 */
const CERT_FIXTURE = fileURLToPath(new URL('./fixtures/certEncrypted.pdf', import.meta.url))
/** structurally broken (PDFBox PDFBOX-6041) → FPDF error 3 */
const CORRUPT_FIXTURE = fileURLToPath(new URL('./fixtures/corruptExample.pdf', import.meta.url))

describe('convertWithPasswordRetry (pure state machine)', () => {
  it('returns the result without prompting when no password is needed', async () => {
    const prompt = vi.fn()
    const result = await convertWithPasswordRetry(() => Promise.resolve('ok'), prompt)
    expect(result).toBe('ok')
    expect(prompt).not.toHaveBeenCalled()
  })

  it('loops wrong passwords through the retry prompt until one works', async () => {
    const attempts: Array<string | undefined> = []
    const convert = (password: string | undefined): Promise<string> => {
      attempts.push(password)
      if (password !== 'sesame') return Promise.reject(passwordRequired())
      return Promise.resolve('opened')
    }
    const prompts: boolean[] = []
    const answers = ['wrong', 'also-wrong', 'sesame']
    const result = await convertWithPasswordRetry(convert, (retry) => {
      prompts.push(retry)
      return Promise.resolve(answers.shift() ?? null)
    })
    expect(result).toBe('opened')
    expect(attempts).toEqual([undefined, 'wrong', 'also-wrong', 'sesame'])
    // first prompt is the plain variant, rejected submissions re-prompt with retry
    expect(prompts).toEqual([false, true, true])
  })

  it('resolves null when the user cancels the prompt', async () => {
    const result = await convertWithPasswordRetry(
      () => Promise.reject(passwordRequired()),
      () => Promise.resolve(null),
    )
    expect(result).toBeNull()
  })

  it('cancel after a wrong attempt also resolves null', async () => {
    const answers: Array<string | null> = ['wrong', null]
    const result = await convertWithPasswordRetry(
      () => Promise.reject(passwordRequired()),
      () => Promise.resolve(answers.shift() ?? null),
    )
    expect(result).toBeNull()
  })

  it('rethrows non-password load errors without prompting', async () => {
    const prompt = vi.fn()
    await expect(
      convertWithPasswordRetry(() => Promise.reject(new PdfLoadError('corrupt', 3)), prompt),
    ).rejects.toMatchObject({ code: 'corrupt' })
    expect(prompt).not.toHaveBeenCalled()
  })

  it('rethrows unsupported-security errors without prompting (P24 C)', async () => {
    // FPDF error 5 = certificate-based / unsupported security handler: a hard
    // PDFium boundary — the password dialog must NOT open (no password can
    // ever succeed, the user would try in vain)
    const prompt = vi.fn()
    await expect(
      convertWithPasswordRetry(() => Promise.reject(new PdfLoadError('unsupported', 5)), prompt),
    ).rejects.toMatchObject({ code: 'unsupported', pdfiumError: 5 })
    expect(prompt).not.toHaveBeenCalled()
  })

  it('rethrows unrelated errors without prompting', async () => {
    const prompt = vi.fn()
    await expect(
      convertWithPasswordRetry(() => Promise.reject(new Error('disk on fire')), prompt),
    ).rejects.toThrow('disk on fire')
    expect(prompt).not.toHaveBeenCalled()
  })
})

describe('integration: testPassword4Spaces.pdf through the real converter', () => {
  it('fails closed without a password', async () => {
    await expect(convertPdfFileToDocxLocal(FIXTURE)).rejects.toMatchObject({
      code: 'password-required',
    })
  })

  it('wrong password → retry prompt → correct password produces a docx', async () => {
    const prompts: boolean[] = []
    const answers = ['wrong', '    ']
    const result = await convertPdfFileToDocxLocalWithPrompt(FIXTURE, (retry) => {
      prompts.push(retry)
      return Promise.resolve(answers.shift() ?? null)
    })
    expect(prompts).toEqual([false, true])
    expect(result).not.toBeNull()
    expect(result!.pages).toBeGreaterThan(0)
    expect(result!.docx.length).toBeGreaterThan(0)
    // docx zip magic
    expect(Array.from(result!.docx.slice(0, 2))).toEqual([0x50, 0x4b])
  }, 60000)

  it('cancelling the prompt aborts silently with null', async () => {
    const result = await convertPdfFileToDocxLocalWithPrompt(FIXTURE, () => Promise.resolve(null))
    expect(result).toBeNull()
  }, 60000)

  it('a certificate-encrypted PDF classifies unsupported and never prompts (P24 C)', async () => {
    const prompt = vi.fn()
    await expect(convertPdfFileToDocxLocalWithPrompt(CERT_FIXTURE, prompt)).rejects.toMatchObject({
      code: 'unsupported',
      pdfiumError: 5,
    })
    expect(prompt).not.toHaveBeenCalled()
  }, 60000)

  it('a corrupt PDF classifies corrupt and never prompts (P24 C)', async () => {
    const prompt = vi.fn()
    await expect(
      convertPdfFileToDocxLocalWithPrompt(CORRUPT_FIXTURE, prompt),
    ).rejects.toMatchObject({ code: 'corrupt', pdfiumError: 3 })
    expect(prompt).not.toHaveBeenCalled()
  }, 60000)
})

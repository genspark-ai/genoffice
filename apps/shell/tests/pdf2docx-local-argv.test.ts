import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

/**
 * Windows file associations launch the packaged app with the document path
 * as process.argv[1]; emscripten copies that into its synthetic environ with
 * an ASCII-only assert, so a CJK path aborted every local PDF conversion with
 * "Aborted(Assertion failed)". Runs in its own file so this module's pdfium
 * singleton initialises under the poisoned argv.
 */
const FIXTURE = fileURLToPath(new URL('./fixtures/testPassword4Spaces.pdf', import.meta.url))

describe('ensurePdfium under a non-ASCII process.argv[1]', () => {
  it('converts to docx and pptx when argv[1] is a CJK document path', async () => {
    process.argv[1] = 'C:\\Users\\\u7528\u6237\\Desktop\\\u62a5\u544a.pdf'
    const { convertPdfFileToDocxLocal } = await import('../src/main/pdf2docx-local')
    const { convertPdfFileToPptxLocal } = await import('../src/main/pdf2pptx-local')
    const docx = await convertPdfFileToDocxLocal(FIXTURE, undefined, '    ')
    expect(docx.pages).toBeGreaterThan(0)
    const pptx = await convertPdfFileToPptxLocal(FIXTURE, undefined, '    ')
    expect(pptx.pages).toBe(docx.pages)
  })
})

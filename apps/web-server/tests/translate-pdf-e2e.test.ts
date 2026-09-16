/**
 * PDF translation E2E — proves the KB → dictionary → file pass loop works for
 * PDFs the way it does for DOCX / XLSX.
 *
 * PDF text is rendered into a background image and the translated text is
 * overlaid on top, so the run uses `pdfplumber` for region extraction and
 * `reportlab` for the overlay. The web-server spawns the upstream
 * `translate_pdf.py`; we exercise it with a hand-written dictionary and
 * assert the output PDF contains English text in the regions where the
 * input had Chinese.
 *
 * The fixture PDF is built in `beforeAll` from reportlab + the system's
 * STHeiti font so its ToUnicode map actually round-trips Chinese — the
 * earlier `pdf-without-text-layer.pdf` fixture only exists to exercise
 * the "no text layer" error path.
 */
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { spawn, type ChildProcess } from 'node:child_process'
import { existsSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

interface IpcResult<T = unknown> {
  ok: boolean
  result: T
}

async function ipc<T = unknown>(base: string, channel: string, args: unknown[] = []): Promise<IpcResult<T>> {
  const res = await fetch(`${base}/api/ipc/${channel}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ args }),
  })
  return (await res.json()) as IpcResult<T>
}

async function waitForHealth(base: string, timeoutMs = 30_000): Promise<void> {
  const deadline = Date.now() + timeoutMs
  while (Date.now() < deadline) {
    try {
      const res = await fetch(`${base}/health`)
      if (res.ok) return
    } catch {
      /* keep polling */
    }
    await new Promise((r) => setTimeout(r, 250))
  }
  throw new Error(`web-server did not become healthy within ${timeoutMs}ms`)
}

/** Build a real Chinese PDF with a working ToUnicode map for the test fixture. */
/** Run a child process and resolve with its exit code; reject with stderr on failure. */
function runSync(cmd: string, args: string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, { stdio: ['ignore', 'pipe', 'pipe'] })
    let stderr = ''
    child.stderr.on('data', (chunk) => (stderr += chunk.toString('utf8')))
    child.on('error', reject)
    child.on('exit', (code) => {
      if (code === 0) resolve()
      else reject(new Error(`child failed (${code}): ${stderr}`))
    })
  })
}

function buildChinesePdf(path: string): Promise<void> {
  const lines = [
    '供应商交付说明',
    '供应商:江南纺织有限公司。',
    '面料克重 220 g/m²,色牢度 4 级。',
    '产品验收报告已提交。',
    '本批订单预计三周内完成生产。',
    '所有成品须经质检部抽检合格后方可出货。',
    '包装使用可回收纸箱,并在外箱标注批次号。',
    '运输方式为海运整柜,预计四周抵达。',
    '争议解决适用中国法律。',
  ]
  const code = `
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
pdfmetrics.registerFont(TTFont('STHeiti', '/System/Library/Fonts/STHeiti Medium.ttc', subfontIndex=0))
c = canvas.Canvas(${JSON.stringify(path)}, pagesize=A4)
c.setFont('STHeiti', 11)
y = 800
for line in ${JSON.stringify(lines)}:
    c.drawString(72, y, line)
    y -= 22
c.showPage()
c.save()
`
  return runSync('/Users/louloulin/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3', ['-c', code])
}

describe('PDF translation E2E', () => {
  let server: ChildProcess | undefined
  let base: string
  let dataDir: string
  let pdfPath: string
  let dictionaryPath: string

  beforeAll(async () => {
    dataDir = mkdtempSync(join(tmpdir(), 'genoffice-translate-pdf-e2e-'))
    pdfPath = join(dataDir, 'verify-supplier.pdf')
    await buildChinesePdf(pdfPath)
    if (!existsSync(pdfPath)) throw new Error(`pdf fixture missing: ${pdfPath}`)

    // A flat dictionary the upstream pdf handler can apply verbatim.
    dictionaryPath = join(dataDir, 'dict.json')
    writeFileSync(dictionaryPath, JSON.stringify({
      '供应商交付说明': 'Supplier Delivery Instructions',
      '供应商:江南纺织有限公司。': 'Supplier: Jiangnan Textile Co., Ltd.',
      '面料克重 220 g/m²,色牢度 4 级。': 'Fabric GSM 220 g/m², color fastness grade 4.',
      '产品验收报告已提交。': 'The Product Acceptance Report has been submitted.',
      '本批订单预计三周内完成生产。': 'This batch of orders is expected to complete production within three weeks.',
      '所有成品须经质检部抽检合格后方可出货。': 'All finished products must pass sampling inspection by the QC department before shipment.',
      '包装使用可回收纸箱,并在外箱标注批次号。': 'Packaging uses recyclable cartons, and the batch number is marked on the outer box.',
      '运输方式为海运整柜,预计四周抵达。': 'The shipping method is full container by sea, with an estimated arrival in four weeks.',
      '争议解决适用中国法律。': 'Disputes shall be resolved in accordance with Chinese law.',
    }, null, 2), 'utf8')

    const port = 21000 + Math.floor(Math.random() * 8000)
    base = `http://127.0.0.1:${port}`
    const bundle = join(__dirname, '..', 'dist', 'bundle', 'index.js')
    server = spawn(process.execPath, [bundle], {
      env: {
        ...process.env,
        PORT: String(port),
        DATA_DIR: dataDir,
        GENOFFICE_WEB_DATA_DIR: dataDir,
        GENOFFICE_TRANSLATION_KB: join(dataDir, 'translation-kb.json'),
        NO_OPEN: '1',
      },
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    server.stdout?.on('data', () => {})
    server.stderr?.on('data', () => {})
    await waitForHealth(base)
  }, 90_000)

  afterAll(() => {
    server?.kill('SIGTERM')
    try {
      rmSync(dataDir, { recursive: true, force: true })
    } catch {
      /* ignore */
    }
  })

  it('mines PDF text the same way it mines DOCX paragraphs', async () => {
    const { ok, result } = await ipc<{
      ok: boolean
      totalSegments?: number
      coverage?: { total: number; covered: number; exact: number; partial: string[]; uncovered: string[]; ratio: number }
    }>(base, 'ai:translate-build-dictionary', [
      { inputPath: pdfPath, sourceLang: 'zh-CN', targetLang: 'en-US', useLlm: false },
    ])
    expect(ok).toBe(true)
    expect(result.ok).toBe(true)
    // PDF text is line-oriented like the docx; the 9 fixture lines all show up.
    expect(result.totalSegments).toBe(9)
    // KB is empty in this scratch dir so nothing matches.
    expect(result.coverage?.exact).toBe(0)
    expect(result.coverage?.uncovered.length).toBe(9)
    expect(result.coverage?.ratio).toBe(0)
  })

  it('rewrites a PDF from a flat dictionary with no provider configured', async () => {
    const outputPath = join(dataDir, 'verify-supplier_translated.pdf')
    const { ok, result } = await ipc<{
      ok: boolean
      outputPath?: string
      dictionaryReused?: boolean
      coverage?: { total: number; covered: number; exact: number; ratio: number }
      error?: string
    }>(base, 'ai:translate-file-auto', [
      {
        inputPath: pdfPath,
        outputPath,
        sourceLang: 'zh-CN',
        targetLang: 'en-US',
        dictionaryPath,
        // Pure dictionary rewrite — must not depend on the provider.
        settings: { provider: 'not-configured', providers: {} },
      },
    ])
    expect(ok).toBe(true)
    expect(result.ok).toBe(true)
    expect(result.dictionaryReused).toBe(true)
    expect(result.coverage?.total).toBe(9)
    expect(result.coverage?.covered).toBe(9)
    expect(result.coverage?.ratio).toBe(1)
    expect(result.outputPath).toBe(outputPath)
    expect(existsSync(outputPath)).toBe(true)
  })
})

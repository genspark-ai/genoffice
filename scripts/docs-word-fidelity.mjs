/**
 * docs vs Word open-consistency comparison test (pre-release regression).
 *
 * Three comparison layers, each building on the last:
 *   L1 page count   — total pages of the Word-exported PDF vs our exported PDF
 *   L2 text flow    — per-page first-line match rate (page breaks) + full-text bigram
 *                     similarity (parse loss)
 *   L3 visual       — rasterize both PDFs page by page → normalize size → pixelmatch,
 *                     reporting whole-page and ink-weighted similarity; low-scoring pages
 *                     get a three-panel comparison image
 *
 * Usage: node scripts/docs-word-fidelity.mjs [--filter substring] [--dir directory]
 *        [--engine lo|word] (default lo) [--refresh-baseline] [--skip-baseline] (cache only)
 * Baseline engines:
 *   lo   — LibreOffice headless, no permissions needed, CI-friendly (has systematic small
 *          layout differences vs Word; good as a regression gate: a sudden score drop
 *          means a parse/render regression on our side)
 *   word — Word for Mac AppleScript, the fidelity gold standard; must run in a terminal
 *          granted automation + file permissions (mkdir hangs silently when the sandbox
 *          container directory lacks TCC authorization)
 * Deps: poppler (pdftotext/pdftoppm), apps/docs/out from npm run build
 * Output: .task/word-fidelity-report/report.md / report.json / diff-*.png
 * Cache: .task/word-fidelity/word-pdf/ keyed by docx mtime+size; unchanged corpus skips Word
 */

import { execFileSync, spawn } from 'node:child_process'
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs'
import { basename, join, resolve } from 'node:path'
import { homedir, tmpdir } from 'node:os'
import WebSocket from 'ws'
import pixelmatch from 'pixelmatch'
import { PNG } from 'pngjs'

const ROOT = resolve(import.meta.dirname, '..')
const ELECTRON = join(
  ROOT,
  'node_modules',
  'electron',
  'dist',
  'Electron.app',
  'Contents',
  'MacOS',
  'Electron',
)
const DOCS_MAIN = join(ROOT, 'apps', 'docs', 'out', 'main', 'index.js')
const CACHE_DIR = join(ROOT, '.task', 'word-fidelity')
const OURS_PDF_DIR = join(CACHE_DIR, 'ours-pdf')
const MANIFEST = join(CACHE_DIR, 'manifest.json')
const REPORT_DIR = join(ROOT, '.task', 'word-fidelity-report')
const WORD_SANDBOX_TMP = join(homedir(), 'Library/Containers/com.microsoft.Word/Data/tmp/fidelity')

const argv = process.argv.slice(2)
const argOf = (name, dflt) => {
  const i = argv.indexOf(name)
  return i >= 0 && argv[i + 1] ? argv[i + 1] : dflt
}
const DOC_DIR = resolve(ROOT, argOf('--dir', 'example/docx'))
const FILTER = argOf('--filter', '')
const ENGINE = argOf('--engine', 'lo')
const REFRESH = argv.includes('--refresh-baseline')
const SKIP_BASELINE = argv.includes('--skip-baseline')
const SKIP_OURS = argv.includes('--skip-ours') // reuse exported PDFs and recompute metrics only (for tuning metric definitions)
const BASE_PDF_DIR = join(CACHE_DIR, `${ENGINE}-pdf`)

const DPI = 72 // rasterization resolution
const CMP_W = 360 // normalized comparison width (downsampling itself acts as blur tolerance)
const PM_THRESHOLD = 0.25
const DIFF_SHOT_BELOW = 0.9 // emit a three-panel image when ink similarity falls below this
// verdict gates (first-round calibration values; can tighten as fidelity improves)
const GATE = {
  pageDeltaWarn: 1,
  textSimFail: 0.85,
  charSimFail: 0.93,
  gridSimWarn: 0.75,
  gridSimFail: 0.5,
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms))
const norm = (s) => s.replace(/\s+/g, '')
// Text comparison normalization: NFKC splits ligature presentation forms; strip punctuation/direction controls/whitespace — squashes RTL extraction noise
const normText = (s) => s.normalize('NFKC').replace(/[\p{P}\p{Cf}\p{Mn}\s]/gu, '')

// ---------- LibreOffice baseline (headless, no permissions needed) ----------

const LO_PROFILE = join(tmpdir(), `lo-fidelity-profile-${process.pid}`)

function loConvert(docxPath, pdfPath) {
  if (existsSync(pdfPath)) rmSync(pdfPath)
  const outDir = resolve(pdfPath, '..')
  execFileSync(
    'soffice',
    [
      '--headless',
      '--norestore',
      `-env:UserInstallation=file://${LO_PROFILE}`,
      '--convert-to',
      'pdf',
      '--outdir',
      outDir,
      docxPath,
    ],
    { timeout: 120_000 },
  )
  const produced = join(outDir, basename(docxPath).replace(/\.docx$/i, '.pdf'))
  if (!existsSync(produced)) throw new Error('PDF was not produced')
  if (produced !== pdfPath) renameSync(produced, pdfPath)
}

// ---------- Word baseline (AppleScript, same approach as pagination-baseline-word.mjs) ----------

function osascript(script) {
  return execFileSync('osascript', ['-e', script], { timeout: 120_000, encoding: 'utf8' }).trim()
}

function restartWord() {
  try {
    osascript('tell application "Microsoft Word" to quit saving no')
  } catch {}
  execFileSync('sleep', ['3'])
  osascript('tell application "Microsoft Word" to launch')
  execFileSync('sleep', ['2'])
}

function recoverWord() {
  try {
    osascript(`
tell application "System Events" to tell process "Microsoft Word"
  if exists window "Grant File Access" then click button "Cancel" of window "Grant File Access"
end tell`)
  } catch {}
  try {
    osascript('tell application "Microsoft Word" to close every document saving no')
  } catch {}
}

function wordConvert(docxPath, pdfPath) {
  if (existsSync(pdfPath)) rmSync(pdfPath)
  const sandboxPdf = join(WORD_SANDBOX_TMP, basename(pdfPath))
  if (existsSync(sandboxPdf)) rmSync(sandboxPdf)
  const script = `
with timeout of 90 seconds
tell application "Microsoft Word"
  open (POSIX file "${docxPath}" as text)
  repeat 150 times
    if (count of documents) > 0 then exit repeat
    delay 0.2
  end repeat
  if (count of documents) is 0 then error "document did not open in 30s"
  set theDoc to document 1
  save as theDoc file name (POSIX file "${sandboxPdf}" as text) file format format PDF
  close theDoc saving no
end tell
end timeout`
  try {
    osascript(script)
  } catch (err) {
    recoverWord()
    throw err
  }
  if (!existsSync(sandboxPdf)) throw new Error('PDF was not produced')
  renameSync(sandboxPdf, pdfPath)
}

// ---------- Our export (CDP-driven, same approach as docs-open-smoke.mjs) ----------

class CDP {
  static async connect(port, retries = 60) {
    for (let i = 0; i < retries; i++) {
      try {
        const tabs = await (await fetch(`http://localhost:${port}/json`)).json()
        const page = tabs.find((t) => t.type === 'page' && t.webSocketDebuggerUrl)
        if (page) {
          const c = new CDP()
          await c._init(page.webSocketDebuggerUrl)
          return c
        }
      } catch {}
      await sleep(500)
    }
    throw new Error('CDP connect timeout')
  }
  _init(url) {
    this.id = 1
    this._pending = new Map()
    return new Promise((res, rej) => {
      this._socket = new WebSocket(url)
      this._socket.onmessage = (ev) => {
        const m = JSON.parse(ev.data.toString())
        if (m.id && this._pending.has(m.id)) {
          const p = this._pending.get(m.id)
          this._pending.delete(m.id)
          if (m.error) p.reject(new Error(m.error.message))
          else p.resolve(m.result)
        }
      }
      this._socket.onerror = () => rej(new Error('ws error'))
      this._socket.onopen = () => res()
    })
  }
  send(method, params = {}) {
    return new Promise((resolve, reject) => {
      const id = this.id++
      this._pending.set(id, { resolve, reject })
      this._socket.send(JSON.stringify({ id, method, params }))
    })
  }
  async evaluate(expression) {
    const r = await this.send('Runtime.evaluate', {
      expression,
      awaitPromise: true,
      returnByValue: true,
    })
    if (r.exceptionDetails)
      throw new Error(r.exceptionDetails.exception?.description ?? r.exceptionDetails.text)
    return r.result?.value
  }
  close() {
    try {
      this._socket?.close()
    } catch {}
  }
}

async function oursExport(docxPath, pdfPath, index) {
  const port = 19680 + (index % 40)
  const stamp = `${process.pid}-${index}`
  const userData = join(tmpdir(), `genoffice-fidelity-${stamp}`)
  mkdirSync(userData, { recursive: true })
  const docCopy = join(tmpdir(), `fidelity-${stamp}-${basename(docxPath)}`)
  copyFileSync(docxPath, docCopy)
  if (existsSync(pdfPath)) rmSync(pdfPath)

  const proc = spawn(
    ELECTRON,
    [
      DOCS_MAIN,
      docCopy,
      '--no-sandbox',
      `--remote-debugging-port=${port}`,
      `--user-data-dir=${userData}`,
    ],
    {
      env: { ...process.env, NODE_ENV: 'production', GENOFFICE_LANG: 'zh' },
      stdio: ['ignore', 'pipe', 'pipe'],
    },
  )

  let cdp
  try {
    cdp = await CDP.connect(port)
    const t0 = Date.now()
    let stable = 0
    let lastPages = -1
    while (Date.now() - t0 < 60_000) {
      let p = null
      try {
        // the CJK literal matches the zh default UI status-bar label for "open failed"
        p = await cdp.evaluate(`(() => {
          const s = document.querySelector('.status-msg')?.textContent ?? ''
          return { failed: s.includes('打开失败'), pages: window.__pageDebug?.slices?.length ?? 0 }
        })()`)
      } catch {}
      if (p?.failed) throw new Error('document failed to open')
      if (p && p.pages >= 1) {
        stable = p.pages === lastPages ? stable + 1 : 0
        lastPages = p.pages
        if (stable >= 2) break
      }
      await sleep(700)
    }
    if (lastPages < 1) throw new Error('no pagination output within 60s')
    await cdp.evaluate('document.fonts.ready.then(() => true)')
    await sleep(500)
    // Open paginated preview before exporting: on the preview path each .pv-page is exactly one
    // sheet (WYSIWYG), whereas the continuous-flow path lets Chromium reflow, making page count
    // and headers/footers diverge from what's on screen
    // the CJK literals match the zh default UI labels for View / Paginated Preview
    const opened = await cdp.evaluate(`(async () => {
      ;[...document.querySelectorAll('button')].find((b) => b.textContent?.trim() === '视图')?.click()
      await new Promise((r) => setTimeout(r, 300))
      const btn = [...document.querySelectorAll('button')].find((b) => b.textContent?.includes('分页预览'))
      if (!btn) return false
      btn.click()
      return true
    })()`)
    if (!opened) throw new Error('paginated-preview button not found')
    let pvPages = 0
    for (let k = 0; k < 60; k++) {
      await sleep(700)
      const n = await cdp.evaluate(`document.querySelectorAll('.pv-page').length`)
      if (n >= 1 && n === pvPages) break
      pvPages = n
    }
    if (pvPages < 1) throw new Error('paginated preview produced no pages')
    lastPages = pvPages
    await cdp.evaluate(`window.__exportPdf(${JSON.stringify(pdfPath)})`)
    for (let i = 0; i < 60 && !existsSync(pdfPath); i++) await sleep(500)
    if (!existsSync(pdfPath)) {
      const s = await cdp.evaluate(`document.querySelector('.status-msg')?.textContent ?? ''`)
      throw new Error(`export produced no PDF (status: ${s.trim()})`)
    }
    return lastPages
  } finally {
    cdp?.close()
    proc.kill('SIGKILL')
    await sleep(300)
    rmSync(userData, { recursive: true, force: true })
    rmSync(docCopy, { force: true })
  }
}

// ---------- Comparison ----------

function pdfPageCount(pdfPath) {
  const out = execFileSync('pdfinfo', [pdfPath], { timeout: 30_000, encoding: 'utf8' })
  return Number(out.match(/^Pages:\s+(\d+)/m)?.[1] ?? 0)
}

/** Monotonic page alignment (Needleman-Wunsch on per-page text bigram similarity) — still pairs identical-content pages after page-break divergence */
function alignPages(wp, op) {
  const n = wp.length
  const m = op.length
  const GAP = -0.05
  const sim = Array.from({ length: n }, (_, i) => op.map((o) => bigramSim(wp[i], o)))
  const dp = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0))
  for (let i = 0; i <= n; i++) dp[i][0] = i * GAP
  for (let j = 0; j <= m; j++) dp[0][j] = j * GAP
  for (let i = 1; i <= n; i++) {
    for (let j = 1; j <= m; j++) {
      dp[i][j] = Math.max(
        dp[i - 1][j - 1] + sim[i - 1][j - 1],
        dp[i - 1][j] + GAP,
        dp[i][j - 1] + GAP,
      )
    }
  }
  const pairs = []
  let i = n
  let j = m
  while (i > 0 && j > 0) {
    if (dp[i][j] === dp[i - 1][j - 1] + sim[i - 1][j - 1]) {
      if (sim[i - 1][j - 1] >= 0.3) pairs.push([i - 1, j - 1])
      i--
      j--
    } else if (dp[i][j] === dp[i - 1][j] + GAP) i--
    else j--
  }
  return pairs.reverse()
}

/** Per-page text, strictly aligned to the pdfinfo page count (image/blank pages kept as empty strings so raster page numbers match) */
function pageTexts(pdfPath, count) {
  const raw = execFileSync('pdftotext', ['-layout', '-enc', 'UTF-8', pdfPath, '-'], {
    timeout: 60_000,
    encoding: 'utf8',
    maxBuffer: 64 * 1024 * 1024,
  })
  const pages = raw.split('\f').slice(0, count)
  while (pages.length < count) pages.push('')
  return pages
}

/** Strip header/footer lines: lines repeated on >=60% of pages are chrome, which inflates
 *  text-metric false positives as page counts diverge and skews first-line matching
 *  (every page's first line becomes the same header) */
function stripChrome(pages) {
  if (pages.length < 4) return pages
  const counts = new Map()
  for (const p of pages) {
    for (const line of new Set(
      p
        .split('\n')
        .map((l) => norm(l))
        .filter(Boolean),
    )) {
      counts.set(line, (counts.get(line) ?? 0) + 1)
    }
  }
  const chrome = new Set(
    [...counts].filter(([, n]) => n >= Math.max(3, pages.length * 0.6)).map(([l]) => l),
  )
  // Page-number line variants (e.g. "Seite N von M") normalize to distinct strings; cluster by prefix and strip separately
  return pages.map((p) =>
    p
      .split('\n')
      .filter((l) => !chrome.has(norm(l)))
      .join('\n'),
  )
}

function firstLine(pageText) {
  for (const line of pageText.split('\n')) {
    const t = line.trim()
    if (t) return t.slice(0, 40)
  }
  return ''
}

function bigramSim(a, b) {
  const grams = (s) => {
    const set = new Set()
    for (let i = 0; i < s.length - 1; i++) set.add(s.slice(i, i + 2))
    return set
  }
  const A = grams(norm(a))
  const B = grams(norm(b))
  if (A.size === 0 && B.size === 0) return 1
  let inter = 0
  for (const g of A) if (B.has(g)) inter++
  return inter / Math.max(A.size + B.size - inter, 1)
}

/** Character multiset Jaccard (unordered). RTL extraction noise only scrambles order without dropping chars: ordered bigram scores low while this stays high; real content loss lowers both */
function charSim(a, b) {
  const count = (s) => {
    const m = new Map()
    for (const ch of normText(s)) m.set(ch, (m.get(ch) ?? 0) + 1)
    return m
  }
  const A = count(a)
  const B = count(b)
  let inter = 0
  let union = 0
  for (const [ch, n] of A) {
    inter += Math.min(n, B.get(ch) ?? 0)
    union += Math.max(n, B.get(ch) ?? 0)
  }
  for (const [ch, n] of B) if (!A.has(ch)) union += n
  return union > 0 ? inter / union : 1
}

function rasterize(pdfPath, outPrefix) {
  execFileSync('pdftoppm', ['-png', '-r', String(DPI), pdfPath, outPrefix], { timeout: 120_000 })
  const dir = resolve(outPrefix, '..')
  const base = basename(outPrefix)
  return readdirSync(dir)
    .filter((f) => f.startsWith(base + '-') && f.endsWith('.png'))
    .sort((a, b) => Number(a.match(/-(\d+)\.png$/)[1]) - Number(b.match(/-(\d+)\.png$/)[1]))
    .map((f) => join(dir, f))
}

/** Box-filter resample into a tw×th RGBA buffer */
function resample(png, tw, th) {
  const { width: sw, height: sh, data } = png
  const out = Buffer.alloc(tw * th * 4)
  for (let y = 0; y < th; y++) {
    const y0 = Math.floor((y * sh) / th)
    const y1 = Math.max(y0 + 1, Math.floor(((y + 1) * sh) / th))
    for (let x = 0; x < tw; x++) {
      const x0 = Math.floor((x * sw) / tw)
      const x1 = Math.max(x0 + 1, Math.floor(((x + 1) * sw) / tw))
      let r = 0,
        g = 0,
        b = 0,
        n = 0
      for (let yy = y0; yy < y1; yy++) {
        for (let xx = x0; xx < x1; xx++) {
          const i = (yy * sw + xx) * 4
          r += data[i]
          g += data[i + 1]
          b += data[i + 2]
          n++
        }
      }
      const o = (y * tw + x) * 4
      out[o] = r / n
      out[o + 1] = g / n
      out[o + 2] = b / n
      out[o + 3] = 255
    }
  }
  return out
}

const inkCount = (buf) => {
  let n = 0
  for (let i = 0; i < buf.length; i += 4) {
    if (0.299 * buf[i] + 0.587 * buf[i + 1] + 0.114 * buf[i + 2] < 245) n++
  }
  return n
}

/** Row/column ink projections */
function inkProfiles(buf, tw, th) {
  const rows = new Float64Array(th)
  const cols = new Float64Array(tw)
  for (let y = 0; y < th; y++) {
    for (let x = 0; x < tw; x++) {
      const i = (y * tw + x) * 4
      const ink = Math.max(0, 245 - (0.299 * buf[i] + 0.587 * buf[i + 1] + 0.114 * buf[i + 2]))
      rows[y] += ink
      cols[x] += ink
    }
  }
  return { rows, cols }
}

/** Optimal 1-D projection shift (minimizes L1 difference) */
function bestShift(a, b, maxShift) {
  let best = 0
  let bestErr = Infinity
  for (let s = -maxShift; s <= maxShift; s++) {
    let err = 0
    for (let i = 0; i < a.length; i++) {
      const j = i + s
      err += Math.abs(a[i] - (j >= 0 && j < b.length ? b[j] : 0))
    }
    if (err < bestErr) {
      bestErr = err
      best = s
    }
  }
  return best
}

/** Shift an RGBA buffer (vacated pixels filled white) */
function shiftBuf(buf, tw, th, dx, dy) {
  const out = Buffer.alloc(buf.length, 255)
  for (let y = 0; y < th; y++) {
    const sy = y + dy
    if (sy < 0 || sy >= th) continue
    for (let x = 0; x < tw; x++) {
      const sx = x + dx
      if (sx < 0 || sx >= tw) continue
      const d = (y * tw + x) * 4
      const s = (sy * tw + sx) * 4
      out[d] = buf[s]
      out[d + 1] = buf[s + 1]
      out[d + 2] = buf[s + 2]
      out[d + 3] = 255
    }
  }
  return out
}

/**
 * Compare our page po vs baseline page pw: stitch baseline pages pw±1 into a strip,
 * use row ink projections to find the best landing window for our page inside the strip,
 * then compare — page-break differences (content shifted across pages) no longer
 * pollute the render fidelity score.
 */
function comparePage(wImgs, pw, oursPng) {
  const o = PNG.sync.read(readFileSync(oursPng))
  const tw = CMP_W
  const th = Math.round((CMP_W * o.height) / o.width)
  let bo = resample(o, tw, th)

  const stripPages = []
  let directOff = 0
  for (let p = pw - 1; p <= pw + 1; p++) {
    if (!wImgs[p]) continue
    if (p < pw) directOff = th
    stripPages.push(resample(PNG.sync.read(readFileSync(wImgs[p])), tw, th))
  }
  const stripH = th * stripPages.length
  const strip = Buffer.concat(stripPages)
  const stripRows = inkProfiles(strip, tw, stripH).rows
  const oRows = inkProfiles(bo, tw, th).rows
  let bestOff = 0
  let bestErr = Infinity
  for (let s = 0; s <= stripH - th; s++) {
    let err = 0
    for (let i = 0; i < th; i++) err += Math.abs(oRows[i] - stripRows[s + i])
    if (err < bestErr) {
      bestErr = err
      bestOff = s
    }
  }
  // Evaluate both alignment hypotheses (best projection window / directly the same page) and keep the more explanatory one
  const windowAt = (off) => strip.subarray(off * tw * 4, (off + th) * tw * 4)
  let bw = windowAt(bestOff)
  if (bestOff !== directOff && gridSim(windowAt(directOff), bo, tw, th) > gridSim(bw, bo, tw, th)) {
    bw = windowAt(directOff)
  }

  // Horizontal registration (systematic page-margin difference)
  const dx = bestShift(
    inkProfiles(bw, tw, th).cols,
    inkProfiles(bo, tw, th).cols,
    Math.round(tw * 0.06),
  )
  if (dx !== 0) bo = shiftBuf(bo, tw, th, dx, 0)

  const diff = new PNG({ width: tw, height: th })
  const diffN = pixelmatch(bw, bo, diff.data, tw, th, { threshold: PM_THRESHOLD, includeAA: false })
  const ink = Math.max(inkCount(bw), inkCount(bo), 1)
  return {
    rawSim: 1 - diffN / (tw * th),
    inkSim: Math.max(0, 1 - diffN / ink),
    gridSim: gridSim(bw, bo, tw, th),
    tw,
    th,
    bw,
    bo,
    diff,
  }
}

/** Grid ink-distribution similarity: coarse cells (>=3x line height) accumulate ink mass;
 *  both sides are normalized and scored as 1 − total variation. Normalization cancels systematic
 *  font-weight/grayscale differences and coarse cells tolerate line-level drift; sensitive to
 *  block-level loss, misplacement, and lost styling — the primary cross-renderer visual metric */
function gridSim(a, b, tw, th, cell = 36) {
  const gw = Math.ceil(tw / cell)
  const gh = Math.ceil(th / cell)
  const da = new Float64Array(gw * gh)
  const db = new Float64Array(gw * gh)
  for (const [buf, d] of [
    [a, da],
    [b, db],
  ]) {
    for (let y = 0; y < th; y++) {
      for (let x = 0; x < tw; x++) {
        const i = (y * tw + x) * 4
        const ink = Math.max(0, 245 - (0.299 * buf[i] + 0.587 * buf[i + 1] + 0.114 * buf[i + 2]))
        d[Math.floor(y / cell) * gw + Math.floor(x / cell)] += ink
      }
    }
  }
  const sa = da.reduce((s, v) => s + v, 0)
  const sb = db.reduce((s, v) => s + v, 0)
  // Both sides nearly blank (mean ink < 0.5/245): treat as identical, so a leftover page-number position difference isn't normalized into a 0 score
  const blank = 0.5 * tw * th
  if (sa < blank && sb < blank) return 1
  if (sa === 0 || sb === 0) return 0
  let tv = 0
  for (let i = 0; i < da.length; i++) tv += Math.abs(da[i] / sa - db[i] / sb)
  return 1 - tv / 2
}

/** Word | ours | diff three-panel image */
function writeTriptych(cmp, outPath) {
  const gap = 4
  const { tw, th } = cmp
  const png = new PNG({ width: tw * 3 + gap * 2, height: th })
  png.data.fill(255)
  const blit = (src, ox) => {
    for (let y = 0; y < th; y++) {
      for (let x = 0; x < tw; x++) {
        const s = (y * tw + x) * 4
        const d = (y * png.width + ox + x) * 4
        png.data[d] = src[s]
        png.data[d + 1] = src[s + 1]
        png.data[d + 2] = src[s + 2]
        png.data[d + 3] = 255
      }
    }
  }
  blit(cmp.bw, 0)
  blit(cmp.bo, tw + gap)
  blit(cmp.diff.data, (tw + gap) * 2)
  writeFileSync(outPath, PNG.sync.write(png))
}

// ---------- Main flow ----------

const files = readdirSync(DOC_DIR)
  .filter((f) => /\.docx$/i.test(f) && !f.startsWith('~$') && !f.startsWith('.'))
  .filter((f) => !FILTER || f.includes(FILTER))
  .sort()
if (files.length === 0) {
  console.error(`no docx found: ${DOC_DIR} filter=${FILTER}`)
  process.exit(2)
}
for (const d of [BASE_PDF_DIR, OURS_PDF_DIR, REPORT_DIR]) mkdirSync(d, { recursive: true })
if (ENGINE === 'word' && !SKIP_BASELINE) mkdirSync(WORD_SANDBOX_TMP, { recursive: true })
const manifest = existsSync(MANIFEST) ? JSON.parse(readFileSync(MANIFEST, 'utf8')) : {}

let baseVersion = `${ENGINE} (cached)`
if (!SKIP_BASELINE) {
  try {
    baseVersion =
      ENGINE === 'word'
        ? `Microsoft Word ${osascript('version of application "Microsoft Word"')}`
        : execFileSync('soffice', ['--version'], { timeout: 60_000, encoding: 'utf8' })
            .split('\n')[0]
            .trim()
  } catch (e) {
    console.error(`baseline engine ${ENGINE} unavailable:`, e.message)
    process.exit(1)
  }
}
console.log(
  `docs ↔ ${ENGINE === 'word' ? 'Word' : 'LibreOffice'} consistency comparison — ${files.length} documents, baseline ${baseVersion}\n`,
)

const results = []
for (let i = 0; i < files.length; i++) {
  const file = files[i]
  const docxPath = join(DOC_DIR, file)
  const stem = basename(file, '.docx')
  const wordPdf = join(BASE_PDF_DIR, `${stem}.pdf`)
  const oursPdf = join(OURS_PDF_DIR, `${stem}.pdf`)
  const st = statSync(docxPath)
  const key = `${st.mtimeMs}:${st.size}`
  const mkey = `${ENGINE}:${stem}`
  const r = { file, verdict: 'FAIL', notes: [] }
  results.push(r)
  process.stdout.write(`[${i + 1}/${files.length}] ${file}\n`)

  // L0a baseline (cached)
  try {
    if (REFRESH || manifest[mkey] !== key || !existsSync(wordPdf)) {
      if (SKIP_BASELINE) throw new Error('--skip-baseline set but no valid cached baseline')
      process.stdout.write(`  ${ENGINE}→pdf ... `)
      if (ENGINE === 'word') {
        try {
          wordConvert(docxPath, wordPdf)
        } catch (first) {
          process.stdout.write(`retry (${String(first.message).split('\n').pop()}) ... `)
          restartWord()
          wordConvert(docxPath, wordPdf)
        }
      } else {
        loConvert(docxPath, wordPdf)
      }
      manifest[mkey] = key
      writeFileSync(MANIFEST, JSON.stringify(manifest, null, 2))
      console.log('✓')
    }
  } catch (e) {
    r.notes.push(`baseline conversion failed: ${e.message}`)
    console.log(`  ❌ ${r.notes.at(-1)}`)
    continue
  }

  // L0b our export
  const oursMetaPath = join(CACHE_DIR, 'ours-meta.json')
  const oursMeta = existsSync(oursMetaPath) ? JSON.parse(readFileSync(oursMetaPath, 'utf8')) : {}
  if (SKIP_OURS && existsSync(oursPdf)) {
    r.displayPages = oursMeta[stem] ?? null
  } else {
    try {
      process.stdout.write('  ours→pdf ... ')
      r.displayPages = await oursExport(docxPath, oursPdf, i)
      oursMeta[stem] = r.displayPages
      writeFileSync(oursMetaPath, JSON.stringify(oursMeta, null, 2))
      console.log('✓')
    } catch (e) {
      r.notes.push(`our export failed: ${e.message}`)
      console.log(`  ❌ ${r.notes.at(-1)}`)
      continue
    }
  }

  // L1 page count + L2 text flow
  r.wordPages = pdfPageCount(wordPdf)
  r.ourPages = pdfPageCount(oursPdf)
  r.pageDelta = r.ourPages - r.wordPages
  const wp = stripChrome(pageTexts(wordPdf, r.wordPages))
  const op = stripChrome(pageTexts(oursPdf, r.ourPages))
  const nCmp = Math.min(wp.length, op.length)
  let startHit = 0
  for (let p = 0; p < nCmp; p++) {
    if (norm(firstLine(wp[p])) === norm(firstLine(op[p]))) startHit++
  }
  r.pageStartMatch = wp.length ? startHit / wp.length : 0
  r.textSim = bigramSim(wp.join('\n'), op.join('\n'))
  r.charSim = charSim(wp.join('\n'), op.join('\n'))

  // L3 visual: monotonically align by page text first, then compare the aligned page pairs (measures render fidelity, decoupled from pagination differences)
  const rasterDir = join(tmpdir(), `fidelity-raster-${process.pid}-${i}`)
  mkdirSync(rasterDir, { recursive: true })
  try {
    const wImgs = rasterize(wordPdf, join(rasterDir, 'w'))
    const oImgs = rasterize(oursPdf, join(rasterDir, 'o'))
    const pairs = alignPages(wp, op)
    const pageSims = []
    for (const [pw, po] of pairs) {
      if (!wImgs[pw] || !oImgs[po]) continue
      const cmp = comparePage(wImgs, pw, oImgs[po])
      pageSims.push({
        wordPage: pw + 1,
        ourPage: po + 1,
        rawSim: cmp.rawSim,
        inkSim: cmp.inkSim,
        gridSim: cmp.gridSim,
      })
      if (cmp.gridSim < DIFF_SHOT_BELOW) {
        writeTriptych(cmp, join(REPORT_DIR, `diff-${stem}-w${pw + 1}-o${po + 1}.png`))
      }
    }
    r.alignedPages = pageSims.length
    r.alignCoverage = r.wordPages ? pageSims.length / r.wordPages : 0
    r.pageSims = pageSims
    const avg = (key) =>
      pageSims.length ? pageSims.reduce((s, p) => s + p[key], 0) / pageSims.length : 0
    r.avgGridSim = avg('gridSim')
    r.minGridSim = pageSims.length ? Math.min(...pageSims.map((p) => p.gridSim)) : 0
    r.avgInkSim = avg('inkSim')
    r.avgRawSim = avg('rawSim')
  } finally {
    rmSync(rasterDir, { recursive: true, force: true })
  }

  // Verdict
  let rtlNoise = false
  if (r.displayPages != null && r.displayPages !== r.ourPages)
    r.notes.push(
      `on-screen ${r.displayPages} pages ≠ exported ${r.ourPages} pages (internal inconsistency)`,
    )
  if (r.textSim < GATE.textSimFail) {
    if (r.charSim < GATE.charSimFail)
      r.notes.push(
        `full-text similarity ${(r.textSim * 100).toFixed(1)}% / char-set ${(r.charSim * 100).toFixed(1)}% too low (suspected parse loss)`,
      )
    else rtlNoise = true // chars intact but reordered: RTL/ligature extraction noise, downgrade to WARN
  }
  if (r.avgGridSim < GATE.gridSimFail)
    r.notes.push(`visual similarity ${(r.avgGridSim * 100).toFixed(1)}% too low`)
  if (r.alignCoverage < 0.5)
    r.notes.push(`only ${(r.alignCoverage * 100).toFixed(0)}% of pages alignable`)
  const warn =
    rtlNoise ||
    Math.abs(r.pageDelta) > GATE.pageDeltaWarn ||
    r.avgGridSim < GATE.gridSimWarn ||
    r.alignCoverage < 0.8
  r.verdict = r.notes.length > 0 ? 'FAIL' : warn ? 'WARN' : 'PASS'
  const mark = r.verdict === 'PASS' ? '✅' : r.verdict === 'WARN' ? '⚠️' : '❌'
  console.log(
    `  ${mark} pages ${r.wordPages}→${r.ourPages} (on-screen ${r.displayPages}) | aligned ${r.alignedPages}/${r.wordPages} | ` +
      `page-start match ${(r.pageStartMatch * 100).toFixed(0)}% | full-text ${(r.textSim * 100).toFixed(1)}% (chars ${(r.charSim * 100).toFixed(1)}%) | ` +
      `visual (grid) avg ${(r.avgGridSim * 100).toFixed(1)}% / min ${(r.minGridSim * 100).toFixed(1)}% ${r.notes.join('; ')}`,
  )
}

// ---------- Report ----------

const passed = results.filter((r) => r.verdict === 'PASS').length
const warned = results.filter((r) => r.verdict === 'WARN').length
const failed = results.filter((r) => r.verdict === 'FAIL').length
const now = new Date().toISOString().replace('T', ' ').slice(0, 19)
const pct = (v) => (v === undefined ? '—' : (v * 100).toFixed(1) + '%')

const md = [
  `# docs ↔ ${ENGINE === 'word' ? 'Word' : 'LibreOffice'} consistency report`,
  ``,
  `- Time: ${now}  Baseline: ${baseVersion} (${DPI}dpi raster, normalized width ${CMP_W}px)`,
  `- Corpus: \`${DOC_DIR}\` (${files.length} documents)`,
  `- Result: **${passed} passed / ${warned} warned / ${failed} failed**`,
  ``,
  `| # | Document | Baseline pages | Our pages (on-screen) | Align rate | Page-start match | Full-text sim | Char-set | Visual avg | Visual min | Verdict | Notes |`,
  `|---|------|--------|--------|--------|----------|----------|--------|----------|----------|------|------|`,
  ...results.map(
    (r, i) =>
      `| ${i + 1} | ${r.file} | ${r.wordPages ?? '—'} | ${r.ourPages ?? '—'}(${r.displayPages ?? '—'}) | ${pct(r.alignCoverage)} | ` +
      `${pct(r.pageStartMatch)} | ${pct(r.textSim)} | ${pct(r.charSim)} | ${pct(r.avgGridSim)} | ${pct(r.minGridSim)} | ${r.verdict} | ${r.notes.join('; ')} |`,
  ),
  ``,
  `Metric definitions: page-start match = share of pages whose break point equals the baseline; full-text sim = bigram Jaccard of both PDFs' extracted text;`,
  `align rate = share of baseline pages pairable after monotonic page-text alignment; visual = grid ink-distribution similarity of aligned page pairs`,
  `after rasterization and shift registration (12px grid, 1 − L1 diff / total ink; robust to font rendering noise, sensitive to block-level loss/misplacement).`,
  `On-screen ≠ exported page count is flagged separately as an internal inconsistency.`,
  `Page pairs with grid similarity below ${DIFF_SHOT_BELOW * 100}% get a triptych \`diff-<doc>-w<baseline page>-o<our page>.png\` (left baseline / middle ours / right diff).`,
  ``,
].join('\n')
writeFileSync(join(REPORT_DIR, 'report.md'), md)
writeFileSync(
  join(REPORT_DIR, 'report.json'),
  JSON.stringify({ time: now, engine: ENGINE, baseVersion, gates: GATE, results }, null, 2),
)

console.log(
  `\n=== ${passed} passed / ${warned} warned / ${failed} failed — report: .task/word-fidelity-report/report.md ===`,
)
process.exit(failed === 0 ? 0 : 1)

/**
 * xlsx rendering fidelity comparison pipeline (sheets counterpart of fidelity-compare.mjs).
 *
 * Reference side: Microsoft Excel for Mac via AppleScript → PDF → pdftoppm PNG (96dpi).
 *   The file is copied into Excel's sandbox container to avoid TCC prompts; osascript runs
 *   under a timeout and Excel is killed on hang (modal repair prompts on corrupt files).
 *   `save as active sheet` still exports the whole workbook, so the comparison scope is
 *   PDF page 1 ↔ the sheet active on open (fine for single-sheet/content-on-sheet-1 samples).
 * Our side: playwright drives the built shell (apps/shell/out) with the xlsx as argv,
 *   waits for the sheets view, composites the Univer canvases and crops off the
 *   row/column headers so the crop matches the print output's content-only framing.
 * Compare: pixelmatch (weak signal here — print layout vs grid viewport never aligns
 *   pixel-perfectly; the HTML report's side-by-side is the primary artifact).
 *
 * Usage: node tools/sheets-fidelity-compare.mjs <a.xlsx> [b.xlsx …] [--out DIR]
 * Prereq: npm run build:all; brew: poppler (pdftoppm); Excel launched manually once.
 */
/* global document, window, PointerEvent */
import { _electron as electron } from 'playwright-core'
import { execFileSync } from 'node:child_process'
import * as fs from 'node:fs'
import * as os from 'node:os'
import * as path from 'node:path'
import { PNG } from 'pngjs'
import pixelmatch from 'pixelmatch'

const ROOT = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..')
const ELECTRON_BIN = path.join(
  ROOT,
  'node_modules/electron/dist/Electron.app/Contents/MacOS/Electron',
)
const SHELL_DIR = path.join(ROOT, 'apps/shell')
const EXCEL_BOX = path.join(
  process.env.HOME,
  'Library/Containers/com.microsoft.Excel/Data/fidelity-tmp',
)

const args = process.argv.slice(2)
const files = []
let outDir = '/tmp/sheets-fidelity/run'
for (let i = 0; i < args.length; i++) {
  if (args[i] === '--out') outDir = args[++i]
  else files.push(path.resolve(args[i]))
}
if (!files.length) {
  console.error('usage: node tools/sheets-fidelity-compare.mjs <a.xlsx> … [--out DIR]')
  process.exit(1)
}
fs.rmSync(outDir, { recursive: true, force: true })
fs.mkdirSync(outDir, { recursive: true })

const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

/** Sheet names in workbook order with visibility; [0] of the visible ones maps to PDF page 1. */
function sheetInfo(xlsx) {
  try {
    const wb = execFileSync('unzip', ['-p', xlsx, 'xl/workbook.xml'], {
      maxBuffer: 64e6,
      stdio: ['ignore', 'pipe', 'ignore'],
    }).toString()
    return [...wb.matchAll(/<sheet [^>]*>/g)].map((m) => ({
      name: m[0]
        .match(/ name="([^"]+)"/)[1]
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"'),
      visible: !/ state="(hidden|veryHidden)"/.test(m[0]),
    }))
  } catch {
    return []
  }
}

/** Excel → PDF → PNGs; returns png paths. Kills Excel on osascript timeout. */
function exportRef(xlsx, dir) {
  fs.mkdirSync(dir, { recursive: true })
  fs.mkdirSync(EXCEL_BOX, { recursive: true })
  const src = path.join(EXCEL_BOX, 'in.xlsx')
  const pdf = path.join(EXCEL_BOX, 'out.pdf')
  fs.copyFileSync(xlsx, src)
  fs.rmSync(pdf, { force: true })
  const script = `
    tell application "Microsoft Excel"
      open (POSIX file "${src}")
      set wb to active workbook
      save workbook as wb filename "${pdf}" file format PDF file format
      close active workbook saving no
    end tell`
  try {
    execFileSync('osascript', ['-e', script], { stdio: 'pipe', timeout: 120_000 })
  } catch (e) {
    try {
      execFileSync('pkill', ['-x', 'Microsoft Excel'], { stdio: 'ignore' })
    } catch {
      /* no process to kill */
    }
    throw new Error('Excel export failed: ' + (e.stderr?.toString() || e.message), { cause: e })
  } finally {
    fs.rmSync(src, { force: true })
  }
  if (!fs.existsSync(pdf)) throw new Error('Excel produced no pdf')
  fs.renameSync(pdf, path.join(dir, 'ref.pdf'))
  execFileSync('pdftoppm', ['-png', '-r', '96', path.join(dir, 'ref.pdf'), path.join(dir, 'ref')], {
    stdio: 'pipe',
  })
  return fs
    .readdirSync(dir)
    .filter((f) => /^ref-\d+\.png$/.test(f))
    .sort((a, b) => parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10))
    .map((f) => path.join(dir, f))
}

/**
 * Launch the shell with the xlsx as argv, wait for the sheets view (footer sheet tabs are
 * DOM — any sheet name in body.textContent means the load finished), composite the Univer
 * canvases and crop the row/col header band (46×24 CSS px at the canvas's own scale).
 */
async function shootOurs(xlsx, dir, sheets) {
  const names = sheets.map((s) => s.name)
  const firstVisible = sheets.find((s) => s.visible)?.name
  fs.mkdirSync(dir, { recursive: true })
  const userDataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'genoffice-fidelity-'))
  const app = await electron.launch({
    executablePath: ELECTRON_BIN,
    args: [SHELL_DIR, xlsx],
    env: {
      ...process.env,
      GENOFFICE_USER_DATA: userDataDir,
      AI_OFFICE_USER_DATA: userDataDir,
      GENOFFICE_LANG: 'en',
    },
    timeout: 30_000,
  })
  try {
    // the sheets editor is a WebContentsView that appears as its own page
    let page
    const deadline = Date.now() + 30_000
    while (!page) {
      for (const w of app.windows()) {
        const href = await w.evaluate(() => window.location.href).catch(() => '')
        if (href.includes('sheets/out')) page = w
      }
      if (page) break
      if (Date.now() > deadline) throw new Error('sheets view never appeared')
      await app.waitForEvent('window', { timeout: 1_000 }).catch(() => {})
    }
    await app.evaluate(({ BrowserWindow }) => {
      BrowserWindow.getAllWindows()[0]?.setBounds({ x: 0, y: 0, width: 1600, height: 1000 })
    })
    for (let i = 0; i < 80; i++) {
      const loaded = await page
        .evaluate(
          (ns) =>
            !!document.querySelector('#univer-container canvas') &&
            (ns.length === 0 || ns.some((n) => document.body.textContent.includes(n))),
          names,
        )
        .catch(() => false)
      if (loaded) break
      await sleep(500)
    }
    // PDF page 1 comes from the first visible sheet, but the file may open on another
    // active sheet — click the matching footer tab so both sides show the same sheet
    if (firstVisible) {
      await page
        .evaluate((n) => {
          const tab = [...document.querySelectorAll('span.univer-truncate')].find(
            (el) => el.textContent.trim() === n,
          )
          if (!tab) return false
          tab
            .closest('[class*="cursor-pointer"]')
            ?.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }))
          tab.click()
          return true
        }, firstVisible)
        .catch(() => false)
      await sleep(1_500)
    }
    await sleep(3_000) // canvas paint, chart/image async render
    // park the selection in the viewport's bottom-right corner so the A1 highlight
    // doesn't sit on top of the content being compared
    {
      const box = await page
        .evaluate(() => {
          const c = [...document.querySelectorAll('#univer-container canvas')].find(
            (x) => x.width > 200 && x.height > 200,
          )
          if (!c) return null
          const r = c.getBoundingClientRect()
          return { x: r.right - 50, y: r.bottom - 50 }
        })
        .catch(() => null)
      if (box) {
        await page.mouse.click(box.x, box.y).catch(() => {})
        await sleep(400)
      }
    }
    const dataUrl = await page.evaluate(() => {
      const container = document.querySelector('#univer-container')
      if (!container) return null
      const canvases = [...container.querySelectorAll('canvas')].filter(
        (c) => c.width > 200 && c.height > 200,
      )
      if (!canvases.length) return null
      const base = canvases[0].getBoundingClientRect()
      const scale = canvases[0].width / base.width // device pixels per CSS px
      const headerW = Math.round(46 * scale)
      const headerH = Math.round(24 * scale)
      const out = document.createElement('canvas')
      out.width = canvases[0].width - headerW
      out.height = canvases[0].height - headerH
      const ctx = out.getContext('2d')
      ctx.fillStyle = '#ffffff'
      ctx.fillRect(0, 0, out.width, out.height)
      for (const c of canvases) {
        const r = c.getBoundingClientRect()
        ctx.drawImage(
          c,
          (r.left - base.left) * scale - headerW,
          (r.top - base.top) * scale - headerH,
          c.width,
          c.height,
        )
      }
      return out.toDataURL('image/png')
    })
    if (!dataUrl) throw new Error('canvas compositing failed')
    const f = path.join(dir, 'ours-1.png')
    fs.writeFileSync(f, Buffer.from(dataUrl.split(',')[1], 'base64'))
    return [f]
  } finally {
    await app.close().catch(() => app.process().kill())
    fs.rmSync(userDataDir, { recursive: true, force: true })
  }
}

/** RGBA bilinear scaling. */
function resize(png, w, h) {
  const out = new PNG({ width: w, height: h })
  for (let y = 0; y < h; y++) {
    const sy = (y * png.height) / h
    const y0 = Math.min(Math.floor(sy), png.height - 1)
    const y1 = Math.min(y0 + 1, png.height - 1)
    const fy = sy - y0
    for (let x = 0; x < w; x++) {
      const sx = (x * png.width) / w
      const x0 = Math.min(Math.floor(sx), png.width - 1)
      const x1 = Math.min(x0 + 1, png.width - 1)
      const fx = sx - x0
      for (let c = 0; c < 4; c++) {
        const p00 = png.data[(y0 * png.width + x0) * 4 + c]
        const p01 = png.data[(y0 * png.width + x1) * 4 + c]
        const p10 = png.data[(y1 * png.width + x0) * 4 + c]
        const p11 = png.data[(y1 * png.width + x1) * 4 + c]
        out.data[(y * w + x) * 4 + c] =
          p00 * (1 - fx) * (1 - fy) + p01 * fx * (1 - fy) + p10 * (1 - fx) * fy + p11 * fx * fy
      }
    }
  }
  return out
}

function diffPair(refPath, oursPath, diffPath) {
  const ref = PNG.sync.read(fs.readFileSync(refPath))
  let ours = PNG.sync.read(fs.readFileSync(oursPath))
  if (ours.width !== ref.width || ours.height !== ref.height)
    ours = resize(ours, ref.width, ref.height)
  const diff = new PNG({ width: ref.width, height: ref.height })
  const bad = pixelmatch(ref.data, ours.data, diff.data, ref.width, ref.height, { threshold: 0.18 })
  fs.writeFileSync(diffPath, PNG.sync.write(diff))
  return bad / (ref.width * ref.height)
}

const rows = []
for (const xlsx of files) {
  const name = path.basename(xlsx).replace(/\.xlsx$/i, '')
  const fileDir = path.join(outDir, name)
  console.log(`\n=== ${name} ===`)
  const sheets = sheetInfo(xlsx)
  let refs
  try {
    refs = exportRef(xlsx, path.join(fileDir, 'ref'))
  } catch (e) {
    console.error('  reference export failed:', e.message.split('\n')[0])
    rows.push({ file: name, error: 'ref: ' + e.message.split('\n')[0] })
    continue
  }
  let ours
  try {
    ours = await shootOurs(xlsx, path.join(fileDir, 'ours'), sheets)
  } catch (e) {
    console.error('  genoffice shot failed:', e.message.split('\n')[0])
    rows.push({ file: name, error: 'ours: ' + e.message.split('\n')[0] })
    continue
  }
  const diffPath = path.join(fileDir, 'diff-1.png')
  const pct = diffPair(refs[0], ours[0], diffPath)
  rows.push({ file: name, pct, pages: refs.length, ref: refs[0], ours: ours[0], diff: diffPath })
  console.log(`  page 1: mismatch ${(pct * 100).toFixed(1)}% (ref pages: ${refs.length})`)
}

const rel = (p) => path.relative(outDir, p)
const html = `<!doctype html><meta charset="utf-8"><title>xlsx fidelity comparison</title>
<style>body{font:13px -apple-system,sans-serif;margin:16px;background:#f6f6f8}
h2{margin:24px 0 8px}table{border-collapse:collapse}td{padding:4px;vertical-align:top;text-align:center}
img{width:420px;border:1px solid #ccc;background:#fff}
.pct{font-weight:600}.bad{color:#c00}.ok{color:#080}.err{color:#c00;font-weight:600}</style>
<h1>xlsx fidelity comparison (reference: Excel)</h1>
${rows
  .map((r) =>
    r.error
      ? `<h2>${r.file} · <span class="err">${r.error}</span></h2>`
      : `
<h2>${r.file} · <span class="pct ${r.pct > 0.2 ? 'bad' : 'ok'}">${(r.pct * 100).toFixed(1)}% mismatch</span> · ${r.pages} ref page(s)</h2>
<table><tr><td>Excel<br><img src="${rel(r.ref)}"></td><td>GenOffice Sheets<br><img src="${rel(r.ours)}"></td><td>diff<br><img src="${rel(r.diff)}"></td></tr></table>`,
  )
  .join('')}
`
fs.writeFileSync(path.join(outDir, 'report.html'), html)
fs.writeFileSync(
  path.join(outDir, 'summary.json'),
  JSON.stringify(
    rows.map(({ file, pct, pages, error }) => ({ file, pct, pages, error })),
    null,
    2,
  ),
)
console.log('\nreport →', path.join(outDir, 'report.html'))
const worst = rows
  .filter((r) => !r.error)
  .sort((a, b) => b.pct - a.pct)
  .slice(0, 10)
console.log('worst files:')
for (const r of worst) console.log(`  ${r.file}: ${(r.pct * 100).toFixed(1)}%`)
console.log('errors:', rows.filter((r) => r.error).length)

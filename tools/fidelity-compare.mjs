/**
 * pptx rendering fidelity comparison pipeline.
 *
 * Reference side: LibreOffice headless (default) or PowerPoint for Mac (--ref powerpoint)
 *   export to PDF → pdftoppm PNG per page (96dpi, 16:9 → 1280×720).
 *   (PowerPoint AppleScript needs its first-run screens clicked through once per machine;
 *   before that the process starts windowless and every apple event times out.)
 * Our side: playwright drives the packaged GenOffice Slides (Electron) with zoom locked at 100%,
 *   clicks through the thumbnails page by page, and screenshots the canvas element .stage-rel.
 * Compare: pixelmatch per-pixel diff (bilinear-scaled to the same size first), emitting a side-by-side HTML report.
 *
 * Usage: node tools/fidelity-compare.mjs <a.pptx> [b.pptx …] [--max-slides N] [--out DIR] [--ref powerpoint]
 * Prereq: npm run build -w @genoffice/slides; brew: libreoffice + poppler (pdftoppm).
 */
/* global document, window, MouseEvent -- used inside page.evaluate() browser context */
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
const APP_DIR = path.join(ROOT, 'apps/slides')

const args = process.argv.slice(2)
const files = []
let maxSlides = 12
let outDir = '/tmp/fidelity/run'
let refKind = 'libreoffice'
let refCache = null
for (let i = 0; i < args.length; i++) {
  if (args[i] === '--max-slides') maxSlides = parseInt(args[++i], 10)
  else if (args[i] === '--out') outDir = args[++i]
  else if (args[i] === '--ref')
    refKind = args[++i] // libreoffice | powerpoint
  else if (args[i] === '--ref-cache')
    refCache = path.resolve(args[++i]) // previous run's outDir: reuse <cache>/<deck>/ref instead of re-exporting
  else files.push(path.resolve(args[i]))
}
if (!files.length) {
  console.error('usage: node tools/fidelity-compare.mjs <a.pptx> … [--max-slides N] [--out DIR]')
  process.exit(1)
}
fs.rmSync(outDir, { recursive: true, force: true })
fs.mkdirSync(outDir, { recursive: true })

const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

/**
 * Thumbnail indexes of visible slides, in presentation order. LibreOffice/PowerPoint skip
 * hidden slides (show="0") when exporting PDF, so the reference page k maps to thumbnail
 * visibleSlideIndexes(pptx)[k], not k.
 */
function visibleSlideIndexes(pptx) {
  const read = (entry) =>
    execFileSync('unzip', ['-p', pptx, entry], { maxBuffer: 64 * 1024 * 1024 }).toString('utf8')
  const rels = read('ppt/_rels/presentation.xml.rels')
  const relMap = {}
  for (const tag of rels.match(/<Relationship\b[^>]*>/g) ?? []) {
    const id = tag.match(/ Id="([^"]+)"/)?.[1]
    const target = tag.match(/ Target="([^"]+)"/)?.[1]
    if (id && target) relMap[id] = target
  }
  const pres = read('ppt/presentation.xml')
  const sldIds = [...pres.matchAll(/<p:sldId\b[^>]*r:id="([^"]+)"/g)].map((m) => m[1])
  const visible = []
  sldIds.forEach((rid, i) => {
    const target = relMap[rid]?.replace(/^\/?(ppt\/)?/, 'ppt/')
    if (!target) return
    const hidden = /<p:sld\b[^>]*\bshow="0"/.test(read(target).slice(0, 4096))
    if (!hidden) visible.push(i)
  })
  return { visible, total: sldIds.length }
}

/** LibreOffice → PDF → PNG per page; returns the array of png paths. */
function exportRef(pptx, dir) {
  fs.mkdirSync(dir, { recursive: true })
  if (refKind === 'powerpoint') {
    exportPdfViaPowerPoint(pptx, path.join(dir, 'ref.pdf'))
  } else {
    execFileSync('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', dir, pptx], {
      stdio: 'pipe',
      timeout: 120_000,
    })
    const pdf = path.join(dir, path.basename(pptx).replace(/\.pptx$/i, '.pdf'))
    if (!fs.existsSync(pdf)) throw new Error('soffice produced no pdf: ' + pdf)
    fs.renameSync(pdf, path.join(dir, 'ref.pdf'))
  }
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
 * PowerPoint for Mac → PDF via AppleScript. The file is copied into PowerPoint's sandbox
 * container so neither open nor save trips a TCC consent prompt. A modal dialog (e.g. repair
 * prompt on a corrupt deck) would hang osascript forever, so on timeout PowerPoint is killed
 * to unblock the remaining decks.
 */
function exportPdfViaPowerPoint(pptx, outPdf) {
  const box = path.join(
    process.env.HOME,
    'Library/Containers/com.microsoft.Powerpoint/Data/fidelity-tmp',
  )
  fs.mkdirSync(box, { recursive: true })
  const src = path.join(box, 'in.pptx')
  const pdf = path.join(box, 'out.pdf')
  fs.copyFileSync(pptx, src)
  fs.rmSync(pdf, { force: true })
  // Stale Office owner-lock files (~$in.pptx, left by a killed save) make every
  // subsequent open fail with -9074 until removed
  for (const f of fs.readdirSync(path.dirname(src))) {
    if (f.startsWith('~$')) fs.rmSync(path.join(path.dirname(src), f), { force: true })
  }
  const script = `
    tell application "Microsoft PowerPoint"
      open (POSIX file "${src}")
      set p to active presentation
      save p in (POSIX file "${pdf}") as save as PDF
      close p saving no
    end tell`
  try {
    execFileSync('osascript', ['-e', script], { stdio: 'pipe', timeout: 120_000 })
  } catch (e) {
    try {
      execFileSync('pkill', ['-x', 'Microsoft PowerPoint'], { stdio: 'ignore' })
    } catch {
      /* no process to kill */
    }
    throw new Error('PowerPoint export failed: ' + (e.message ?? e), { cause: e })
  } finally {
    fs.rmSync(src, { force: true })
  }
  if (!fs.existsSync(pdf)) throw new Error('PowerPoint produced no pdf')
  fs.renameSync(pdf, outPdf)
}

/**
 * Opens GenOffice Slides and, page by page, composites the main canvas's Konva layers into a PNG
 * (in-page toDataURL, unaffected by window size/zoom/DPR; output = slide logical pixels 1280×720).
 */
async function shootOurs(pptx, dir, thumbIndexes) {
  fs.mkdirSync(dir, { recursive: true })
  // Own userData → own single-instance lock: a dev run or the installed app must not kill the eval instance
  const userData =
    process.env.GENOFFICE_USER_DATA ?? path.join(os.tmpdir(), 'genoffice-fidelity-userdata')
  fs.mkdirSync(userData, { recursive: true })
  const app = await electron.launch({
    executablePath: ELECTRON_BIN,
    args: [APP_DIR, pptx],
    env: { ...process.env, GENOFFICE_USER_DATA: userData },
    timeout: 30_000,
  })
  try {
    const page =
      app.windows().find((w) => !w.url().startsWith('devtools://')) ?? (await app.firstWindow())
    // wait for the open to finish (thumbnails appear)
    for (let i = 0; i < 60; i++) {
      const n = await page.evaluate(() => document.querySelectorAll('.thumb').length)
      if (n > 0) break
      await sleep(500)
    }
    const total = await page.evaluate(() => document.querySelectorAll('.thumb').length)
    // Hide the "Click to add title" placeholder hints; the event re-renders the canvas
    // even on single-slide decks where no slide switch follows.
    await page.evaluate(() => {
      window.__genofficeHidePhPrompts = true
      window.dispatchEvent(new Event('genoffice:hide-ph-prompts'))
    })
    await sleep(300)
    const shots = []
    for (const idx of thumbIndexes.filter((i) => i < total)) {
      await page.evaluate((i) => {
        const thumbs = document.querySelectorAll('.thumb')
        thumbs[i]?.dispatchEvent(new MouseEvent('click', { bubbles: true }))
      }, idx)
      await sleep(600) // wait for image decode / repaint
      // Office-private FontFaces (doc-fonts.ts) may still be loading on first pages:
      // wait for the renderer's sync flag (bounded), then for the face loads it started
      for (let i = 0; i < 20; i++) {
        if (await page.evaluate(() => window.__genofficeDocFontsSynced !== false)) break
        await sleep(150)
      }
      await page.evaluate(() => document.fonts.ready).catch(() => {})
      const dataUrl = await page.evaluate(() => {
        const stage = document.querySelector('.stage-rel')
        if (!stage) return null
        const canvases = [...stage.querySelectorAll('canvas')]
        if (!canvases.length) return null
        const w = stage.clientWidth
        const h = stage.clientHeight
        const out = document.createElement('canvas')
        out.width = w
        out.height = h
        const ctx = out.getContext('2d')
        ctx.fillStyle = '#ffffff'
        ctx.fillRect(0, 0, w, h)
        // Each canvas is larger than the slide (CANVAS_BLEED ring for the page shadow) and
        // CSS-transformed by zoom; place it by its on-screen box mapped into slide-logical
        // coordinates so the output is exactly the slide rect at widthPx × heightPx.
        const sr = stage.getBoundingClientRect()
        for (const c of canvases) {
          const r = c.getBoundingClientRect()
          ctx.drawImage(
            c,
            ((r.left - sr.left) * w) / sr.width,
            ((r.top - sr.top) * h) / sr.height,
            (r.width * w) / sr.width,
            (r.height * h) / sr.height,
          )
        }
        return out.toDataURL('image/png')
      })
      if (!dataUrl) throw new Error('canvas compositing failed (thumb ' + (idx + 1) + ')')
      const f = path.join(dir, `ours-${shots.length + 1}.png`)
      fs.writeFileSync(f, Buffer.from(dataUrl.split(',')[1], 'base64'))
      shots.push(f)
    }
    return shots
  } finally {
    await app.close().catch(() => {})
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
for (const pptx of files) {
  const name = path.basename(pptx).replace(/\.pptx$/i, '')
  const deckDir = path.join(outDir, name)
  console.log(`\n=== ${name} ===`)
  let refs
  try {
    const cached = refCache && path.join(refCache, name, 'ref')
    if (
      cached &&
      fs.existsSync(cached) &&
      fs.readdirSync(cached).some((f) => /^ref-\d+\.png$/.test(f))
    ) {
      fs.cpSync(cached, path.join(deckDir, 'ref'), { recursive: true })
      refs = fs
        .readdirSync(path.join(deckDir, 'ref'))
        .filter((f) => /^ref-\d+\.png$/.test(f))
        .sort((a, b) => parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10))
        .map((f) => path.join(deckDir, 'ref', f))
      console.log(`  reusing ${refs.length} cached reference page(s)`)
    } else {
      refs = exportRef(pptx, path.join(deckDir, 'ref'))
    }
  } catch (e) {
    console.error('reference export failed:', e.message)
    continue
  }
  let visible
  try {
    const scan = visibleSlideIndexes(pptx)
    // LibreOffice omits hidden slides from the PDF; PowerPoint may include them. Trust the
    // page count: a full-length PDF means no mapping is needed.
    visible = refs.length >= scan.total ? refs.map((_, i) => i) : scan.visible
    if (visible.length < scan.total)
      console.log(
        `  ${scan.total - visible.length} hidden slide(s) skipped to match the PDF export`,
      )
  } catch (e) {
    console.error('  visible-slide scan failed, assuming none hidden:', e.message)
    visible = refs.map((_, i) => i)
  }
  const n = Math.min(refs.length, visible.length, maxSlides)
  const ours = await shootOurs(pptx, path.join(deckDir, 'ours'), visible.slice(0, n))
  for (let i = 0; i < Math.min(n, ours.length); i++) {
    const diffPath = path.join(deckDir, `diff-${i + 1}.png`)
    const pct = diffPair(refs[i], ours[i], diffPath)
    rows.push({ deck: name, slide: i + 1, pct, ref: refs[i], ours: ours[i], diff: diffPath })
    console.log(`  slide ${i + 1}: mismatch ${(pct * 100).toFixed(1)}%`)
  }
}

// HTML report
const rel = (p) => path.relative(outDir, p)
const html = `<!doctype html><meta charset="utf-8"><title>pptx fidelity comparison</title>
<style>body{font:13px -apple-system,sans-serif;margin:16px;background:#f6f6f8}
h2{margin:24px 0 8px}table{border-collapse:collapse}td{padding:4px;vertical-align:top;text-align:center}
img{width:420px;border:1px solid #ccc;background:#fff}
.pct{font-weight:600}.bad{color:#c00}.ok{color:#080}</style>
<h1>pptx fidelity comparison (reference: ${refKind})</h1>
${rows
  .map(
    (r) => `
<h2>${r.deck} · slide ${r.slide} · <span class="pct ${r.pct > 0.08 ? 'bad' : 'ok'}">${(r.pct * 100).toFixed(1)}% mismatch</span></h2>
<table><tr><td>reference<br><img src="${rel(r.ref)}"></td><td>GenOffice Slides<br><img src="${rel(r.ours)}"></td><td>diff<br><img src="${rel(r.diff)}"></td></tr></table>`,
  )
  .join('')}
`
fs.writeFileSync(path.join(outDir, 'report.html'), html)
console.log('\nreport →', path.join(outDir, 'report.html'))
const worst = [...rows].sort((a, b) => b.pct - a.pct).slice(0, 8)
console.log('worst slides:')
for (const r of worst) console.log(`  ${r.deck} #${r.slide}: ${(r.pct * 100).toFixed(1)}%`)

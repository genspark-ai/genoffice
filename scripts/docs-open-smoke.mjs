/**
 * Batch open smoke test for docs (pre-release regression).
 * Opens every docx under example/docx one by one in the real docs app and asserts:
 *   open succeeds (no "open failed" in the status bar) / pagination yields >=1 page /
 *   body is non-empty / no uncaught exceptions.
 * Archives a screenshot per document and emits Markdown + JSON reports.
 *
 * Usage: node scripts/docs-open-smoke.mjs [--filter substring] [--dir directory] [--timeout ms]
 * Prereq: npm run build (electron-vite output must match the current code).
 * The default --dir (example/docx) is gitignored and not shipped — supply
 * your own corpus of .docx files there or point --dir at one.
 * Output: .task/docs-open-report/report.md / report.json / *.png
 */

import { spawn } from 'node:child_process'
import { copyFileSync, mkdirSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { basename, join, resolve } from 'node:path'
import { tmpdir } from 'node:os'
import WebSocket from 'ws'

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
const REPORT_DIR = join(ROOT, '.task', 'docs-open-report')
const sleep = (ms) => new Promise((r) => setTimeout(r, ms))

const argv = process.argv.slice(2)
const argOf = (name, dflt) => {
  const i = argv.indexOf(name)
  return i >= 0 && argv[i + 1] ? argv[i + 1] : dflt
}
const DOC_DIR = resolve(ROOT, argOf('--dir', 'example/docx'))
const FILTER = argOf('--filter', '')
const TIMEOUT = Number(argOf('--timeout', '60000'))
const BASE_PORT = 19620

class CDP {
  static async connect(port, onEvent, retries = 60) {
    for (let i = 0; i < retries; i++) {
      try {
        const tabs = await (await fetch(`http://localhost:${port}/json`)).json()
        const page = tabs.find((t) => t.type === 'page' && t.webSocketDebuggerUrl)
        if (page) {
          const c = new CDP()
          await c._init(page.webSocketDebuggerUrl, onEvent)
          return c
        }
      } catch {}
      await sleep(500)
    }
    throw new Error('CDP connect timeout')
  }
  _init(url, onEvent) {
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
        } else if (m.method) {
          onEvent?.(m.method, m.params)
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
  async screenshot(path) {
    const r = await this.send('Page.captureScreenshot', { format: 'png' })
    writeFileSync(path, Buffer.from(r.data, 'base64'))
  }
  close() {
    try {
      this._socket?.close()
    } catch {}
  }
}

// the CJK literal below matches the zh default UI status-bar label for "open failed"
const PROBE = `(() => {
  const status = document.querySelector('.status-msg')?.textContent ?? ''
  const dbg = window.__pageDebug
  const pm = document.querySelector('.ProseMirror')
  return {
    ready: !!pm,
    status,
    openFailed: status.includes('打开失败'),
    pages: dbg?.slices?.length ?? 0,
    blocks: dbg?.blocks?.length ?? 0,
    chars: pm ? pm.innerText.replace(/\\s+/g, '').length : 0,
    tables: document.querySelectorAll('.doc-table').length,
    images: document.querySelectorAll('.ProseMirror img').length,
    brokenImages: [...document.querySelectorAll('.ProseMirror img')]
      .filter((i) => i.complete && i.naturalWidth === 0).length,
  }
})()`

async function testOne(file, index) {
  const name = basename(file)
  const port = BASE_PORT + (index % 40)
  const stamp = `${process.pid}-${index}`
  const userData = join(tmpdir(), `genoffice-docsmoke-${stamp}`)
  mkdirSync(userData, { recursive: true })
  // Open a copy in tmp so any app-side write-back (recent list aside) never touches the original corpus
  const docCopy = join(tmpdir(), `docsmoke-${stamp}-${name}`)
  copyFileSync(file, docCopy)

  const errors = []
  const t0 = Date.now()
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

  const result = {
    file: name,
    sizeKB: Math.round(statSync(file).size / 1024),
    verdict: 'FAIL',
    pages: 0,
    blocks: 0,
    chars: 0,
    tables: 0,
    images: 0,
    loadMs: 0,
    notes: [],
    errors,
  }
  let cdp
  try {
    cdp = await CDP.connect(port, (method, params) => {
      if (method === 'Runtime.exceptionThrown') {
        errors.push(
          params.exceptionDetails?.exception?.description ??
            params.exceptionDetails?.text ??
            'exception',
        )
      } else if (method === 'Log.entryAdded' && params.entry?.level === 'error') {
        errors.push(params.entry.text)
      }
    })
    await cdp.send('Runtime.enable')
    await cdp.send('Log.enable')

    // Poll until pagination yields pages and the count stays stable for two rounds, or the open fails / times out
    let probe = null
    let stable = 0
    let lastPages = -1
    while (Date.now() - t0 < TIMEOUT) {
      try {
        probe = await cdp.evaluate(PROBE)
      } catch {
        probe = null
      }
      if (probe?.openFailed) break
      if (probe?.ready && probe.pages >= 1) {
        stable = probe.pages === lastPages ? stable + 1 : 0
        lastPages = probe.pages
        if (stable >= 2) break
      }
      await sleep(700)
    }
    result.loadMs = Date.now() - t0
    if (probe)
      Object.assign(result, {
        pages: probe.pages,
        blocks: probe.blocks,
        chars: probe.chars,
        tables: probe.tables,
        images: probe.images,
      })

    if (!probe?.ready) result.notes.push('timeout: editor not ready')
    else if (probe.openFailed) result.notes.push(`open failed: ${probe.status}`)
    else if (probe.pages < 1) result.notes.push('timeout: no pagination output')
    else {
      const empty = probe.chars === 0 && probe.images === 0
      if (empty) result.notes.push('body is empty (no text, no images)')
      if (probe.brokenImages > 0) result.notes.push(`${probe.brokenImages} images failed to load`)
      const fatal = errors.filter((e) => !/net::|favicon|Autofill/.test(e))
      if (fatal.length > 0) result.notes.push(`${fatal.length} renderer-process errors`)
      result.verdict = empty ? 'FAIL' : result.notes.length > 0 ? 'WARN' : 'PASS'
    }
    const shot = join(
      REPORT_DIR,
      `${String(index + 1).padStart(2, '0')}-${name.replace(/\.docx$/i, '')}.png`,
    )
    try {
      await cdp.screenshot(shot)
      result.screenshot = basename(shot)
    } catch {}
  } catch (e) {
    result.notes.push(`runtime error: ${e.message}`)
  } finally {
    cdp?.close()
    proc.kill('SIGKILL')
    await sleep(300)
    rmSync(userData, { recursive: true, force: true })
    rmSync(docCopy, { force: true })
  }
  return result
}

const files = readdirSync(DOC_DIR)
  .filter((f) => /\.docx$/i.test(f) && !f.startsWith('~$') && !f.startsWith('.'))
  .filter((f) => !FILTER || f.includes(FILTER))
  .sort()
  .map((f) => join(DOC_DIR, f))
if (files.length === 0) {
  console.error(`no docx found: ${DOC_DIR} filter=${FILTER}`)
  process.exit(2)
}
mkdirSync(REPORT_DIR, { recursive: true })

console.log(`docs batch-open smoke test — ${files.length} documents\n`)
const all = []
for (let i = 0; i < files.length; i++) {
  const r = await testOne(files[i], i)
  all.push(r)
  const mark = r.verdict === 'PASS' ? '✅' : r.verdict === 'WARN' ? '⚠️' : '❌'
  console.log(
    `${mark} [${i + 1}/${files.length}] ${r.file} — ${r.pages} pages ${r.chars} chars ${r.loadMs}ms ${r.notes.join('; ')}`,
  )
}

const passed = all.filter((r) => r.verdict === 'PASS').length
const warned = all.filter((r) => r.verdict === 'WARN').length
const failed = all.filter((r) => r.verdict === 'FAIL').length
const now = new Date().toISOString().replace('T', ' ').slice(0, 19)

const md = [
  `# docs batch-open smoke test report`,
  ``,
  `- Time: ${now}`,
  `- Corpus: \`${DOC_DIR}\` (${files.length} documents)`,
  `- Result: **${passed} passed / ${warned} warned / ${failed} failed**`,
  ``,
  `| # | Document | Size | Pages | Chars | Tables | Images | Duration | Verdict | Notes |`,
  `|---|------|------|------|------|------|------|------|------|------|`,
  ...all.map(
    (r, i) =>
      `| ${i + 1} | ${r.file} | ${r.sizeKB}KB | ${r.pages} | ${r.chars} | ${r.tables} | ${r.images} | ${(r.loadMs / 1000).toFixed(1)}s | ${r.verdict} | ${r.notes.join('; ')} |`,
  ),
  ``,
  `Screenshots: \`NN-<doc name>.png\` in the same directory, for manual multilingual glyph/layout checks.`,
  ``,
].join('\n')
writeFileSync(join(REPORT_DIR, 'report.md'), md)
writeFileSync(
  join(REPORT_DIR, 'report.json'),
  JSON.stringify({ time: now, dir: DOC_DIR, passed, warned, failed, results: all }, null, 2),
)

console.log(
  `\n=== ${passed} passed / ${warned} warned / ${failed} failed — report: .task/docs-open-report/report.md ===`,
)
process.exit(failed === 0 ? 0 : 1)

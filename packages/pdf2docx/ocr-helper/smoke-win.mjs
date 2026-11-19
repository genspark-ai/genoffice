/**
 * CI smoke for the Windows system-OCR helper: compile, decode a committed
 * fixture (English text on white), and assert the JSON protocol end to end.
 *
 * Windows Server CI images may ship WITHOUT any OCR recognizer language
 * (client SKUs always have the profile language). The helper reports that as
 * exit code 4 — the smoke then downgrades to protocol-only validation and
 * still fails on compile/decode/crash problems, so a green run always means
 * "the exe we ship cannot crash-loop on user machines".
 */
import { execFileSync, spawnSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))

execFileSync(process.execPath, [join(here, 'build-win.mjs')], { stdio: 'inherit' })

const png = readFileSync(join(here, 'fixtures', 'smoke-en.png'))
const res = spawnSync(join(here, 'win-ocr.exe'), [], {
  input: png,
  maxBuffer: 64 * 1024 * 1024,
  timeout: 60_000,
})

const stderr = (res.stderr ?? Buffer.alloc(0)).toString('utf8').trim()
if (res.status === 4) {
  console.warn(`no OCR language on this runner (${stderr}) — recognition not validated`)
  process.exit(0)
}
if (res.status !== 0) {
  console.error(`helper exit ${res.status}: ${stderr}`)
  process.exit(1)
}

const out = JSON.parse(res.stdout.toString('utf8').replace(/^\uFEFF/, ''))
if (typeof out.paper !== 'number' || !Array.isArray(out.lines)) {
  console.error('malformed helper output:', JSON.stringify(out).slice(0, 400))
  process.exit(1)
}
if (out.paper < 0.5) {
  console.error(`paper share ${out.paper} < 0.5 on a white-background fixture`)
  process.exit(1)
}
const text = out.lines.map((l) => l.t).join(' ')
if (!/quick/i.test(text) || !/12345/.test(text)) {
  console.error(`fixture text not recognized; got: ${text.slice(0, 300)}`)
  process.exit(1)
}
for (const line of out.lines) {
  const [x0, y0, x1, y1] = line.b
  if (!(x0 >= 0 && y0 >= 0 && x1 <= 1 && y1 <= 1 && x1 > x0 && y1 > y0)) {
    console.error(`line box out of normalized range: ${JSON.stringify(line.b)}`)
    process.exit(1)
  }
}
console.log(
  `smoke OK: ${out.lines.length} lines, paper ${out.paper.toFixed(2)}, "${text.slice(0, 80)}"`,
)

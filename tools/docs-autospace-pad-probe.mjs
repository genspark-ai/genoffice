#!/usr/bin/env node
/**
 * Headless-Chromium probe for the CJK/Latin autospace pad (styles.css
 * .doc-autospace-pad). Not Electron and not the corpus gate: it pins the three
 * properties jsdom cannot lay out, under every line-height form the renderer
 * emits (px, calc(... * 1em), unitless):
 *   1. the pad adds no line-box height (a padded line equals its pad-less twin),
 *   2. the seam gap inside a line is Word's whole gap (2 x the 1/8 em base),
 *   3. the gap is dropped when the line breaks at the seam (next line at x = 0).
 *
 *   node tools/docs-autospace-pad-probe.mjs        exit 1 on any failure
 *
 * Needs the playwright chromium build (npx playwright install chromium).
 */
/* global document, getComputedStyle */
import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..')
const { chromium } = createRequire(import.meta.url)('playwright')

const css = readFileSync(join(ROOT, 'apps/docs/src/renderer/styles.css'), 'utf8')
/** the pad rules as committed: @property + .doc-autospace-pad* blocks */
const padCss = [
  ...css.matchAll(/(@property --doc-autospace-gap|\.doc-autospace-pad[\w.:-]*)\s*\{[^}]*\}/g),
]
  .map((m) => m[0])
  .join('\n')

const html = `<!doctype html><meta charset=utf-8><style>
body{font:11pt Calibri,'Helvetica Neue',sans-serif;--doc-autospace-pad:0.125em}
.box{white-space:pre-wrap;margin:8px;width:600px}
.px{line-height:17.9px}.calc{line-height:calc(1.2207 * 1em)}.num{line-height:1.35}
.cjk{font-family:'Hiragino Mincho ProN','Yu Mincho','Noto Serif CJK JP',serif}
#m{position:absolute;white-space:pre;visibility:hidden}
${padCss}
</style>
${['px', 'calc', 'num']
  .map(
    (lh) => `
<div class="box ${lh}" id="${lh}-twin"><span>SNS</span><span class="cjk">\u5e83\u544a\u3001\u30dd\u30b9\u30bf\u30fc\u7b49\uff08\u7a0e\u8fbc\uff09</span></div>
<div class="box ${lh}" id="${lh}-pad"><span>SNS</span><span class="cjk"><span class="doc-autospace-pad">\u5e83</span>\u544a\u3001\u30dd\u30b9\u30bf\u30fc\u7b49\uff08\u7a0e\u8fbc\uff09</span></div>`,
  )
  .join('')}
<span id=m><span>PR TIMES</span><span class=cjk>\u914d\u4fe1</span>/<span class=cjk>\u5236\u4f5c\u3001</span><span>SNS</span></span>
<div id=wrap class="box px"><span>PR TIMES</span><span class=cjk>\u914d\u4fe1</span>/<span class=cjk>\u5236\u4f5c\u3001</span><span>SNS</span><span class="cjk"><span class="doc-autospace-pad">\u5e83</span>\u544a\u3001\u30dd\u30b9\u30bf\u30fc\u7b49\uff08\u7a0e\u8fbc\uff09</span></div>
`

const browser = await chromium.launch()
const page = await browser.newPage()
await page.setContent(html)
const r = await page.evaluate(() => {
  const h = (id) => document.getElementById(id).getBoundingClientRect().height
  const out = { lines: {}, gap: null, wrapStart: null }
  for (const lh of ['px', 'calc', 'num']) out.lines[lh] = [h(`${lh}-twin`), h(`${lh}-pad`)]
  const right = (id) => {
    const el = document.getElementById(id)
    return el.querySelector('.cjk').getBoundingClientRect().right - el.getBoundingClientRect().left
  }
  out.gap = right('px-pad') - right('px-twin')
  const wrap = document.getElementById('wrap')
  wrap.style.width = document.getElementById('m').getBoundingClientRect().width + 6 + 'px'
  const pad = wrap.querySelector('.doc-autospace-pad')
  out.wrapStart = pad.getBoundingClientRect().left - wrap.getBoundingClientRect().left
  out.fontPx = parseFloat(getComputedStyle(pad).fontSize)
  return out
})
await browser.close()

const fails = []
for (const [lh, [twin, pad]] of Object.entries(r.lines))
  if (Math.abs(twin - pad) > 0.01)
    fails.push(`line-height ${lh}: pad line ${pad} px vs twin ${twin} px`)
const want = r.fontPx * 0.25
if (Math.abs(r.gap - want) > 0.1)
  fails.push(`seam gap ${r.gap.toFixed(2)} px, want 0.25 em = ${want.toFixed(2)} px`)
if (Math.abs(r.wrapStart) > 0.1)
  fails.push(`line after a break at the seam starts at ${r.wrapStart.toFixed(2)} px, want 0`)
console.log(JSON.stringify(r))
if (fails.length) {
  for (const f of fails) console.error(`FAIL ${f}`)
  process.exit(1)
}
console.log('autospace pad: no line-box growth, whole gap inside a line, no gap at a line start')

import { test, expect } from '@playwright/test'
import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'

// Use the production stylesheet: jsdom cannot resolve percentage positioning.
const css = readFileSync(resolve('apps/docs/src/renderer/styles.css'), 'utf8')

for (const width of [800, 1100]) {
  for (const zoom of [0.5, 1, 1.5]) {
    for (const boxes of [false, true]) {
      test(`header/footer alignment: width ${width}, zoom ${zoom}, boxes ${boxes}`, async ({
        page,
      }) => {
        const strip = (kind: string, gap = false) =>
          `<div class="page-hf page-hf-${kind}${gap ? ' page-gap-hf' : ''}${boxes ? ' page-hf-has-boxes' : ''}"><span>${kind}</span></div>`
        await page.setContent(`<!doctype html>
          <style>${css}</style>
          <body style="display:block;margin:20px">
            <div style="zoom:${zoom};width:max-content">
              <div class="page-wrap" style="width:max-content;--page-w:${width}px;--section-content-w:${width - 192}px;--markup-w:240px">
                ${strip('header')}${strip('footer')}
                <div class="doc-page" style="box-sizing:border-box;width:${width}px;padding:96px">
                  <div class="page-gap" style="width:${width}px;margin-left:-96px;height:200px">
                    ${strip('header', true)}${strip('footer', true)}
                  </div>
                </div>
              </div>
            </div>
          </body>`)

        for (const side of ['none', 'right', 'left']) {
          const offsets = await page.evaluate(
            ({ side, zoom }) => {
              const wrap = document.querySelector('.page-wrap')!
              wrap.classList.toggle('has-markup-area', side !== 'none')
              wrap.classList.toggle('markup-left', side === 'left')
              const paperLeft = wrap.querySelector('.doc-page')!.getBoundingClientRect().left
              return [...wrap.querySelectorAll('.page-hf span')].map(
                (span) => (span.getBoundingClientRect().left - paperLeft) / zoom,
              )
            },
            { side, zoom },
          )
          expect(offsets).toHaveLength(4)
          for (const offset of offsets) expect(offset, `${side} markup`).toBeCloseTo(96, 1)
        }
      })
    }
  }
}

/**
 * html2docx <input.html> <output.docx> [--verbose]
 *
 * Runs through Vite's SSR loader so the extractor's `?raw` fragment imports
 * resolve the same way they do in the app bundle and in vitest:
 *   npx tsx packages/html2docx/tools/cli.ts in.html out.docx
 *
 * Debug env: H2D_SHOT_TEXT_PATH (write rasterized text), H2D_IR_PATH (dump
 * IR + trace), H2D_TRACE_SELECTOR (in-page probes).
 */
import { writeFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { createServer } from 'vite'
import type { ConvertOptions, ConvertResult } from '../src'
import type { PlaywrightDriver as PlaywrightDriverType } from '../src/drivers/playwright'

const A4 = { width: 794, height: 1123, deviceScaleFactor: 2 }

async function main(): Promise<void> {
  const verbose = process.argv.includes('--verbose') || process.argv.includes('-v')
  const args = process.argv.slice(2).filter((a) => a !== '--verbose' && a !== '-v')
  if (args.length < 2) {
    console.error('Usage: html2docx <input.html> <output.docx> [--verbose]')
    process.exit(1)
  }
  const [input, output] = args
  const root = resolve(dirname(fileURLToPath(import.meta.url)), '..')
  const vite = await createServer({
    root,
    configFile: false,
    logLevel: 'error',
    server: { middlewareMode: true, hmr: false },
    optimizeDeps: { noDiscovery: true },
  })
  try {
    const lib = (await vite.ssrLoadModule('/src/index.ts')) as {
      convertHtmlToDocx: (
        input: { url: string },
        driver: PlaywrightDriverType,
        options: ConvertOptions,
      ) => Promise<ConvertResult>
    }
    const { PlaywrightDriver, launchChrome } = (await vite.ssrLoadModule(
      '/src/drivers/playwright.ts',
    )) as typeof import('../src/drivers/playwright')
    const browser = await launchChrome()
    try {
      const driver = await PlaywrightDriver.create(browser, A4)
      const url = input.startsWith('http') ? input : pathToFileURL(resolve(input)).href
      const result = await lib.convertHtmlToDocx({ url }, driver, {
        log: verbose ? (m) => console.error(m) : undefined,
        traceSelector: process.env.H2D_TRACE_SELECTOR,
        onIr: process.env.H2D_IR_PATH
          ? (ir, trace) => {
              writeFileSync(`${process.env.H2D_IR_PATH}.trace`, JSON.stringify(trace, null, 1))
              writeFileSync(
                process.env.H2D_IR_PATH!,
                JSON.stringify(ir, (key, value) => (key === 'buffer' ? undefined : value), 1),
              )
            }
          : undefined,
      })
      await driver.close()
      writeFileSync(output, result.docx)
      if (process.env.H2D_SHOT_TEXT_PATH) {
        writeFileSync(process.env.H2D_SHOT_TEXT_PATH, result.screenshotText, 'utf8')
      }
      console.log(output)
    } finally {
      await browser.close()
    }
  } finally {
    await vite.close()
  }
}

main().catch((e) => {
  console.error(e)
  process.exit(1)
})

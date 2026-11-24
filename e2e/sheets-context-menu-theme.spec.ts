/** Real Univer portals must use the document's scheme, independently of app chrome. */
import { test, expect } from '@playwright/test'
import { build } from 'esbuild'
import { createServer, type Server } from 'node:http'
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { basename, join, resolve } from 'node:path'

const repo = resolve(__dirname, '..')
let server: Server
let url: string

function luminance(color: string): number {
  const channels = color
    .match(/[\d.]+/g)!
    .slice(0, 3)
    .map(Number)
    .map((value) => {
      const channel = value / 255
      return channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4
    })
  return channels[0] * 0.2126 + channels[1] * 0.7152 + channels[2] * 0.0722
}

function contrast(first: string, second: string): number {
  const values = [luminance(first), luminance(second)].sort((a, b) => a - b)
  return (values[1] + 0.05) / (values[0] + 0.05)
}

test.use({ browserName: 'chromium', headless: true, viewport: { width: 1200, height: 900 } })
test.beforeAll(async ({ browserName }, info) => {
  expect(browserName).toBe('chromium')
  const output = info.outputPath('fixture')
  mkdirSync(output, { recursive: true })
  // Use the production factory, preset, palette, and CSS. Univer itself creates
  // the menu DOM and sets .univer-dark from darkMode; no menu markup is mocked.
  await build({
    stdin: {
      resolveDir: repo,
      loader: 'ts',
      contents: `
        import { createUniver } from ${JSON.stringify(join(repo, 'apps/sheets/src/renderer/create-univer'))};
        import { LocaleType, UniverInstanceType } from '@univerjs/core';
        import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
        import enUS from '@univerjs/preset-sheets-core/locales/en-US';
        import { greenTheme } from '@univerjs/themes';
        import ${JSON.stringify(join(repo, 'packages/ui/src/tokens.css'))};
        import '@univerjs/preset-sheets-core/lib/index.css';
        import ${JSON.stringify(join(repo, 'apps/sheets/src/renderer/styles.css'))};
        const params = new URLSearchParams(location.search);
        const app = params.get('app');
        const doc = params.get('doc');
        if (app !== 'system') document.documentElement.dataset.theme = app;
        document.documentElement.dataset.docTheme = doc;
        const runtime = createUniver({
          theme: greenTheme, darkMode: doc === 'dark', locale: LocaleType.EN_US,
          locales: { [LocaleType.EN_US]: enUS },
          presets: [UniverSheetsCorePreset({ container: 'univer-container',
            header: true, toolbar: false, contextMenu: true })],
        });
        runtime.univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
          id: 'theme-workbook', name: 'Theme test', sheetOrder: ['sheet1'],
          sheets: { sheet1: { id: 'sheet1', name: 'Sheet1', rowCount: 50, columnCount: 20,
            cellData: { 0: { 0: { v: 'Synthetic data' } } } } },
        });
      `,
    },
    bundle: true,
    outfile: join(output, 'bundle.js'),
    loader: { '.png': 'dataurl', '.svg': 'dataurl', '.ttf': 'file', '.woff2': 'file' },
    define: { 'process.env.NODE_ENV': '"production"' },
  })
  server = createServer((request, response) => {
    const path = new URL(request.url ?? '/', 'http://localhost').pathname
    if (path === '/') {
      response.setHeader('Content-Type', 'text/html; charset=utf-8')
      response.end(
        '<!doctype html><link rel="stylesheet" href="/bundle.css">' +
          '<div id="univer-container" style="width:1100px;height:800px"></div>' +
          '<script src="/bundle.js"></script>',
      )
      return
    }
    try {
      response.setHeader(
        'Content-Type',
        path.endsWith('.css')
          ? 'text/css'
          : path.endsWith('.js')
            ? 'application/javascript'
            : 'application/octet-stream',
      )
      response.end(readFileSync(join(output, basename(path))))
    } catch {
      response.writeHead(404).end()
    }
  })
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('missing fixture port')
  url = `http://127.0.0.1:${address.port}`
})
test.afterAll(async () => {
  await new Promise<void>((resolve, reject) =>
    server.close((error) => (error ? reject(error) : resolve())),
  )
})

for (const scenario of [
  { app: 'light', doc: 'light', system: 'light' },
  { app: 'dark', doc: 'light', system: 'dark' },
  { app: 'light', doc: 'dark', system: 'light' },
  { app: 'dark', doc: 'dark', system: 'dark' },
  { app: 'system', doc: 'light', system: 'dark' },
  { app: 'system', doc: 'dark', system: 'light' },
] as const) {
  test(`${scenario.app} app / ${scenario.doc} document / ${scenario.system} system: submenu labels stay readable`, async ({
    page,
  }, info) => {
    await page.emulateMedia({ colorScheme: scenario.system })
    await page.goto(`${url}/?app=${scenario.app}&doc=${scenario.doc}`)
    const grid = page.locator('canvas').filter({ visible: true })
    await expect.poll(async () => grid.count()).toBeGreaterThan(0)
    const gridBounds = () =>
      grid.evaluateAll((canvases) => {
        const rect = canvases
          .map((canvas) => canvas.getBoundingClientRect())
          .find((r) => r.width > 500 && r.height > 300)
        return rect ? { x: rect.x, y: rect.y } : null
      })
    await expect.poll(gridBounds).not.toBeNull()
    const bounds = (await gridBounds())!
    await expect
      .poll(() => page.locator('html').evaluate((html) => html.classList.contains('univer-dark')))
      .toBe(scenario.doc === 'dark')
    await page.mouse.click(bounds!.x + 80, bounds!.y + 70, { button: 'right' })
    const insert = page.getByRole('button', { name: 'Insert', exact: true })
    await expect(insert).toBeVisible()
    const primaryBefore = await insert.evaluate((button) => ({
      color: getComputedStyle(button).color,
      scheme: getComputedStyle(button).colorScheme,
    }))
    await insert.hover()
    const submenu = page.locator('[data-u-context-menu-submenu]')
    await expect(submenu).toBeVisible()
    await expect(submenu.locator('button').first()).toContainText('Insert')
    const probe = await submenu.evaluate((portal) => {
      const panel = portal.firstElementChild!
      return {
        app: document.documentElement.dataset.theme ?? 'system',
        doc: document.documentElement.dataset.docTheme,
        univerDark: document.documentElement.classList.contains('univer-dark'),
        parent: portal.parentElement!.tagName,
        insidePopup: !!portal.closest('.univer-popup'),
        scheme: getComputedStyle(portal).colorScheme,
        background: getComputedStyle(panel).backgroundColor,
        html: portal.outerHTML,
        buttons: Array.from(portal.querySelectorAll('button')).map((button) => ({
          text: button.textContent?.trim(),
          color: getComputedStyle(button).color,
          scheme: getComputedStyle(button).colorScheme,
          width: button.getBoundingClientRect().width,
          height: button.getBoundingClientRect().height,
        })),
      }
    })
    const primaryAfter = await insert.evaluate((button) => ({
      color: getComputedStyle(button).color,
      scheme: getComputedStyle(button).colorScheme,
    }))
    const report = JSON.stringify({ primaryBefore, primaryAfter, ...probe }, null, 2)
    writeFileSync(info.outputPath('portal.json'), report)
    await info.attach('portal', { body: report, contentType: 'application/json' })
    await page.screenshot({ path: info.outputPath('submenu.png') })
    expect(primaryAfter).toEqual(primaryBefore)
    expect(probe.parent).toBe('BODY')
    expect(probe.insidePopup).toBe(false)
    expect(probe.buttons.length).toBeGreaterThan(0)
    for (const button of probe.buttons) {
      expect(button.text?.length).toBeGreaterThan(0)
      expect(button.width).toBeGreaterThan(0)
      expect(button.height).toBeGreaterThan(0)
      expect(contrast(button.color, probe.background)).toBeGreaterThanOrEqual(4.5)
    }
    expect(probe.scheme).toBe(scenario.doc)
    expect(luminance(probe.background) > 0.5).toBe(scenario.doc === 'light')
  })
}

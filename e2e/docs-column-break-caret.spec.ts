/** Native typing must stay after an inserted column break without adding a line. */
import { test, expect } from '@playwright/test'
import { build } from 'esbuild'
import { createServer, type Server } from 'node:http'
import { mkdirSync, readFileSync } from 'node:fs'
import { basename, join, resolve } from 'node:path'

const repo = resolve(__dirname, '..')
let server: Server
let url: string
interface Snapshot {
  nodes: string[]
  paragraphs: number
  height: number
  caretHeight: number
}
interface ProbeWindow extends Window {
  columnBoot(text: string, pos: number, marked: boolean): void
  columnInsert(): void
  columnProbe(): Snapshot
}

test.use({ browserName: 'chromium', headless: true, viewport: { width: 1000, height: 800 } })
test.beforeAll(async ({ browserName }, info) => {
  expect(browserName).toBe('chromium')
  const output = info.outputPath('fixture')
  mkdirSync(output, { recursive: true })
  await build({
    stdin: {
      resolveDir: repo,
      loader: 'ts',
      contents: `
        import { Editor } from '@tiptap/core';
        import { editorExtensions } from ${JSON.stringify(join(repo, 'apps/docs/src/renderer/editor/extensions'))};
        import { insertColumnBreak } from ${JSON.stringify(join(repo, 'apps/docs/src/renderer/editor/page-break'))};
        import ${JSON.stringify(join(repo, 'packages/ui/src/tokens.css'))};
        import ${JSON.stringify(join(repo, 'apps/docs/src/renderer/styles.css'))};
        let editor;
        window.columnBoot = (text, pos, marked) => {
          editor?.destroy();
          editor = new Editor({ element: document.getElementById('editor'), extensions: editorExtensions,
            editorProps: { attributes: { class: 'doc-page' } },
            content: { type: 'doc', content: [{ type: 'docParagraph', content: text ? [{ type: 'text', text,
              ...(marked ? { marks: [{ type: 'docTextStyle', attrs: { font: 'Arial' } }] } : {}) }] : [] }] } });
          editor.chain().focus().setTextSelection(pos).run();
        };
        window.columnInsert = () => insertColumnBreak(editor);
        window.columnProbe = () => {
          const coords = editor.view.coordsAtPos(editor.state.selection.from);
          return { nodes: editor.state.doc.firstChild.content.content.map(n => n.isText ? n.text : n.attrs.colBreak ? '<COL>' : n.type.name),
            paragraphs: editor.state.doc.childCount,
            height: editor.view.dom.firstElementChild.getBoundingClientRect().height,
            caretHeight: coords.bottom - coords.top };
        };
      `,
    },
    bundle: true,
    outfile: join(output, 'bundle.js'),
    loader: { '.png': 'dataurl', '.svg': 'dataurl', '.ttf': 'file', '.woff2': 'file' },
    define: { 'import.meta.env.DEV': 'false' },
  })
  server = createServer((request, response) => {
    const path = new URL(request.url ?? '/', 'http://localhost').pathname
    if (path === '/') {
      response.setHeader('Content-Type', 'text/html; charset=utf-8')
      response.end(
        '<link rel="stylesheet" href="/bundle.css"><div id="editor"></div><script src="/bundle.js"></script>',
      )
      return
    }
    try {
      response.setHeader(
        'Content-Type',
        path.endsWith('.css')
          ? 'text/css; charset=utf-8'
          : path.endsWith('.js')
            ? 'application/javascript; charset=utf-8'
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
  if (server)
    await new Promise<void>((resolve, reject) =>
      server.close((error) => (error ? reject(error) : resolve())),
    )
})

for (const scenario of [
  { name: 'paragraph end', text: 'before', pos: 7 },
  { name: 'paragraph middle', text: 'beforeafter', pos: 7 },
  { name: 'empty paragraph', text: '', pos: 1 },
  { name: 'direct command', text: 'before', pos: 7, direct: true },
  { name: 'moving back across the break', text: 'before', pos: 7, arrows: true },
  { name: 'formatted paragraph', text: 'before', pos: 7, marked: true },
]) {
  test(`types after a column break: ${scenario.name}`, async ({ page }) => {
    await page.addInitScript(() =>
      Object.defineProperty(navigator, 'platform', { get: () => 'Win32' }),
    )
    await page.goto(url)
    await page.waitForFunction(() => 'columnBoot' in window)
    await page.evaluate(
      ({ text, pos, marked }) => (window as unknown as ProbeWindow).columnBoot(text, pos, !!marked),
      scenario,
    )
    await page.evaluate(
      () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
    )
    const before = await page.evaluate(() => (window as unknown as ProbeWindow).columnProbe())
    if (scenario.direct)
      await page.evaluate(() => (window as unknown as ProbeWindow).columnInsert())
    else await page.keyboard.press('Control+Shift+Enter')
    const inserted = await page.evaluate(() => (window as unknown as ProbeWindow).columnProbe())
    if (scenario.arrows) {
      await page.keyboard.press('ArrowLeft')
      await page.keyboard.press('ArrowRight')
    }
    await page.keyboard.type('XYZ')
    const after = await page.evaluate(() => (window as unknown as ProbeWindow).columnProbe())
    const left = scenario.text.slice(0, scenario.pos - 1)
    const right = scenario.text.slice(scenario.pos - 1)
    expect(after.nodes).toEqual([...(left ? [left] : []), '<COL>', `XYZ${right}`])
    expect(after.paragraphs).toBe(1)
    expect(inserted.caretHeight).toBeGreaterThan(0)
    if (scenario.text && !right) expect(inserted.height).toBeCloseTo(before.height, 3)
  })
}

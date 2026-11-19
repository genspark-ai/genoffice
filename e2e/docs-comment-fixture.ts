import type { Page } from '@playwright/test'
import { createServer } from 'node:http'
import { readFile } from 'node:fs/promises'
import { extname, resolve, sep } from 'node:path'
import { launchSdkTestHost, type SdkTestHost } from './sdk-host'

/** Load a real DOCX in the public editor with isolated, in-memory host storage. */
export async function withCommentEditor(
  data: Buffer,
  run: (page: Page) => Promise<void>,
): Promise<void> {
  const root = resolve(__dirname, '../apps/docs/out/embedded')
  const server = createServer(async (request, response) => {
    const path = new URL(request.url ?? '/', 'http://localhost').pathname
    if (path === '/document.docx') {
      response.end(data)
      return
    }
    if (path === '/') {
      response.setHeader('content-type', 'text/html')
      response.end(`<!doctype html><html><head><link rel="stylesheet" href="/sdk/editor.css"></head><body><div id="root"></div><script type="module">
        import {mountDocsEditor,createExternalDocsHost} from '/sdk/editor.js';
        const defaults={getLanguage:'en',getSystemLocale:'en-US',getTheme:'light',getDocumentTheme:'follow',getAuthorName:'Reviewer',getAutoSaveDefault:{on:false,updatedAt:0},getAiPanelPrefs:{side:'left',fontSize:'default',spellcheck:false,openInNewDocs:false,showAiSidebar:false},spellLanguages:{active:[],available:[]},docPasswordIntentRevision:0};
        const services=new Proxy({}, {get(_,name){if(name==='then')return undefined;if(name.startsWith('on'))return ()=>()=>{};return async()=>defaults[name];}});
        const data=await (await fetch('/document.docx')).arrayBuffer();
        const host=createExternalDocsHost('browser',services);
        host.documents={initialDocument:async()=>({id:'comments',name:'Comments.docx',data}),saveDocument:async()=>({ok:false,committed:false,pending:false})};
        window.editor=await mountDocsEditor(document.getElementById('root'),{host});await window.editor.ready;window.ready=true;
      </script></body></html>`)
      return
    }
    try {
      const file = resolve(root, decodeURIComponent(path.slice(5)))
      if (!path.startsWith('/sdk/') || !file.startsWith(root + sep)) {
        response.writeHead(404).end()
        return
      }
      response.setHeader(
        'content-type',
        (
          { '.js': 'text/javascript', '.css': 'text/css', '.woff2': 'font/woff2' } as Record<
            string,
            string
          >
        )[extname(file)] ?? 'application/octet-stream',
      )
      response.end(await readFile(file))
    } catch {
      response.writeHead(404).end()
    }
  })
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('No comment fixture server port')
  let browser: SdkTestHost | undefined
  try {
    browser = await launchSdkTestHost()
    const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } })
    await page.goto(`http://127.0.0.1:${address.port}/`)
    await page.waitForFunction(() => (window as unknown as { ready: boolean }).ready)
    await run(page)
  } finally {
    await browser?.close()
    await new Promise<void>((done, reject) =>
      server.close((error) => (error ? reject(error) : done())),
    )
  }
}

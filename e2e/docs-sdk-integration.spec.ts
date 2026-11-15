// Modified for ForaOffice: verify portable selection setup across SDK test hosts.
import { test, expect } from '@playwright/test'
import { launchSdkTestHost, sdkHostIsElectron, type SdkTestHost } from './sdk-host'
import { createServer } from 'node:http'
import { execFile } from 'node:child_process'
import { promisify } from 'node:util'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, extname, join, resolve, sep } from 'node:path'
import JSZip from 'jszip'
import { createRequire } from 'node:module'
import type { PdfiumModule } from '../packages/pdf2docx/src'
import type { DocsEditorHandle } from '../apps/docs/src/renderer/embedded'
import type { DocsFileMenuAction } from '../apps/docs/src/renderer/sdk/file-menu'
import type { DocsEditorAction } from '../apps/docs/src/shared/editor-action'
import type {
  DocsRibbonAction,
  DocsRibbonConfiguration,
} from '../apps/docs/src/renderer/sdk/ribbon'
import { withPdfDocument } from '../packages/pdf2docx/src/extract/pdfium'
import { renderPageByIndexPng } from '../packages/pdf2docx/src/extract'

const image =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg=='
type FixtureWindow = Window & {
  editor: DocsEditorHandle
  sdkReady: boolean
  setExtensionPanelOpen(open: boolean): void
  fixture: {
    saved: number[]
    fallback: string[]
    readonly: boolean
    cancel: boolean
    menuActions: DocsFileMenuAction[]
    ribbonActions: DocsRibbonAction[]
    agentActions: DocsEditorAction[]
    documentStates: boolean[]
    saveKinds: boolean[]
    closeState?: { autoSave: boolean }
  }
}

async function createPrintRenderer() {
  const require = createRequire(__filename)
  const { init } = await import('@embedpdf/pdfium')
  const wasm = await readFile(require.resolve('@embedpdf/pdfium/pdfium.wasm'))
  const wrapped = await init({
    wasmBinary: wasm.buffer.slice(wasm.byteOffset, wasm.byteOffset + wasm.byteLength),
  })
  const module = (wrapped.pdfium ?? wrapped) as unknown as PdfiumModule & {
    _PDFiumExt_Init(): void
  }
  module._PDFiumExt_Init()
  return (bytes: Uint8Array) =>
    withPdfDocument(module, bytes, (document) =>
      Array.from({ length: module._FPDF_GetPageCount(document) }, (_, index) => {
        const page = renderPageByIndexPng(module, document, index, 1)
        if (!page) throw new Error('Could not render native print output')
        expect(page.data.byteLength).toBeGreaterThan(0)
        const pdfPage = module._FPDF_LoadPage(document, index)
        if (!pdfPage) throw new Error('Could not load native print page')
        let textPage = 0
        try {
          textPage = module._FPDFText_LoadPage(pdfPage)
          if (!textPage) throw new Error('Could not read native print text')
          const text = Array.from({ length: module._FPDFText_CountChars(textPage) }, (_, i) =>
            String.fromCodePoint(module._FPDFText_GetUnicode(textPage, i)),
          ).join('')
          return { png: page.data, text }
        } finally {
          if (textPage) module._FPDFText_ClosePage(textPage)
          module._FPDF_ClosePage(pdfPage)
        }
      }),
    )
}

test('the real Qt public SDK preserves DOCX resources and identical printed content with view annotations', async () => {
  test.skip(sdkHostIsElectron, 'Qt native adapter has a separate integration target')
  test.skip(process.platform !== 'win32', 'Native Qt fixture uses the configured Windows build')
  test.setTimeout(180_000)
  const executable =
    process.env.FORAOFFICE_QT_SDK_TEST ?? 'D:/temp/foraoffice-qt-build/docs-qt-sdk-test.exe'
  const output = await mkdtemp(join(tmpdir(), 'foraoffice-sdk-output-'))
  const cache = await readFile(join(dirname(executable), 'CMakeCache.txt'), 'utf8')
  const qtConfig = /^Qt6_DIR:PATH=(.+)$/m.exec(cache)?.[1].trim()
  if (!qtConfig) throw new Error('The Qt SDK test build has no Qt6_DIR')
  const qtBin = resolve(qtConfig, '../../../bin')
  await promisify(execFile)(executable, ['-o', join(output, 'native-results.txt') + ',txt'], {
    timeout: 170_000,
    windowsHide: true,
    env: {
      ...process.env,
      PATH: qtBin + ';' + process.env.PATH,
      FORAOFFICE_QT_TEST_NODE: process.execPath,
      FORAOFFICE_QT_TEST_PACKAGE: resolve(__dirname, '../apps/docs/out/sdk'),
      FORAOFFICE_QT_TEST_RESULTS: output,
    },
  })
  const archive = await JSZip.loadAsync(await readFile(join(output, 'hosted.docx')))
  const xml = await archive.file('word/document.xml')!.async('string')
  expect(xml).toContain('word')
  expect(xml).not.toContain('sdk-range-decoration')
  expect(xml).not.toContain('sdk-annotation')
  const pictures = Object.values(archive.files).filter(
    (file) => file.name.startsWith('word/media/') && !file.dir,
  )
  expect(pictures.length).toBeGreaterThan(0)
  expect(await pictures[0].async('base64')).toBe(image)
  const render = await createPrintRenderer()
  const decorated = render(await readFile(join(output, 'annotated.pdf')))
  const clean = render(await readFile(join(output, 'clean.pdf')))
  expect(decorated.length).toBeGreaterThan(0)
  expect(
    decorated
      .map((page) => page.text)
      .join('')
      .replace(/\s+/g, ''),
  ).toContain('samesamesame[invalid')
  expect(decorated).toEqual(clean)
})

test('a browser host owns pathless storage while native inspection and advisory annotations keep their semantics', async () => {
  test.setTimeout(120_000)
  const root = resolve(__dirname, '../apps/docs/out/embedded')
  const server = createServer(async (request, response) => {
    const path = new URL(request.url ?? '/', 'http://localhost').pathname
    if (path === '/') {
      response.setHeader('content-type', 'text/html')
      response.end(`<!doctype html><html><head><link rel="stylesheet" href="/sdk/editor.css"></head><body>
        <div id="root"></div><script type="module">
        import {mountDocsEditor,createExternalDocsHost} from '/sdk/editor.js';
        const defaults={getLanguage:'en',getSystemLocale:'en-US',getTheme:'light',getDocumentTheme:'follow',
          getAuthorName:'SDK fixture',getAutoSaveDefault:{on:false,updatedAt:0},
          getAiPanelPrefs:{side:'left',fontSize:'default',spellcheck:false,openInNewDocs:false,showAiSidebar:false},
          spellLanguages:{active:[],available:[]},docPasswordIntentRevision:0};
        window.fixture={saved:[],fallback:[],readonly:false,cancel:false,menuActions:[],ribbonActions:[],agentActions:[],documentStates:[],saveKinds:[]};
        const events=new Map();
        window.checkClose=()=>{for(const listener of events.get('onCloseCheck')??[])listener();return window.fixture.closeState;};
        const services=new Proxy({}, {get(_,name){
          if(name==='then')return undefined;
          if(name.startsWith('on'))return listener=>{const entries=events.get(name)??new Set();entries.add(listener);events.set(name,entries);return ()=>entries.delete(listener);};
          if(name==='reportCloseCheck')return state=>{window.fixture.closeState=state;};
          if(name==='confirmDocumentReplace')return async()=>true;
          if(name==='readImageSource')return async()=>({base64:'${image}',mime:'image/png',name:'fixture.png'});
          return async(...args)=>{if(name.startsWith('saveDocx'))window.fixture.fallback.push(name);return defaults[name];};
        }});
        const host=createExternalDocsHost('${sdkHostIsElectron ? 'electron' : 'browser'}',services,{submit:async action=>{window.fixture.agentActions.push(action);}});
        const doc=()=>({id:'browser:object',name:'Browser.docx',readOnly:window.fixture.readonly,
          ...(window.fixture.saved.length?{data:Uint8Array.from(window.fixture.saved).buffer}:{})});
        host.documents={autoSave:{enabled:true,intervalMs:600000,onBlur:true},initialDocument:async()=>location.search.includes('empty')?null:doc(),
          newDocument:async()=>window.fixture.cancel?null:doc(),openDocument:async()=>doc(),
          saveDocument:async r=>{window.fixture.saved=Array.from(new Uint8Array(r.data));window.fixture.saveKinds.push(r.auto);
            return {ok:true,committed:true,pending:false,document:{id:'browser:object',name:'Browser.docx'}};}};
        host.checkers={fixture:s=>{const at=s.text.indexOf('[');return at<0?[]:[{from:s.positions[at],to:s.positions[at]+1,message:'Fixture diagnostic'}];}};
        window.editor=await mountDocsEditor(document.getElementById('root'),{host});
        await window.editor.ready;window.sdkReady=true;
        window.remountSdk=async()=>{
          const old=window.editor.documents;
          let finishOpen;
          host.documents.openDocument=()=>new Promise(resolve=>{finishOpen=resolve;});
          const pendingOpen=old.open();
          window.editor.dispose();
          const replacement=createExternalDocsHost('${sdkHostIsElectron ? 'electron' : 'browser'}',services);
          window.fixture.newHostCalls=0;
          replacement.documents={...host.documents,newDocument:async()=>{window.fixture.newHostCalls++;return doc();}};
          window.editor=await mountDocsEditor(document.getElementById('root'),{host:replacement});
          await window.editor.ready;
          finishOpen(doc());
          const oldOpenResult=await pendingOpen;
          let refused=false;try{await old.new();}catch{refused=true;}
          return {refused,calls:window.fixture.newHostCalls,oldOpenResult};
        };
        </script></body></html>`)
      return
    }
    try {
      if (!path.startsWith('/sdk/')) {
        response.writeHead(404).end()
        return
      }
      const file = resolve(root, decodeURIComponent(path.slice(5)))
      if (!file.startsWith(root + sep)) {
        response.writeHead(403).end()
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
  if (!address || typeof address === 'string') throw new Error('No browser fixture port')
  let browser: SdkTestHost | undefined
  const errors: string[] = []
  try {
    browser = await launchSdkTestHost()
    const platforms = sdkHostIsElectron ? [process.platform] : ['MacIntel', 'Win32', 'Linux x86_64']
    for (const platform of platforms) {
      const layoutPage = await browser.newPage({ viewport: { width: 1360, height: 900 } })
      try {
        if (!sdkHostIsElectron) {
          await layoutPage.addInitScript((platform) => {
            Object.defineProperty(navigator, 'platform', { get: () => platform })
          }, platform)
        }
        await layoutPage.goto(`http://127.0.0.1:${address.port}/`)
        await layoutPage.waitForFunction(() => (window as unknown as FixtureWindow).sdkReady)
        for (const theme of ['light', 'dark']) {
          await layoutPage.evaluate(
            (theme) => document.documentElement.setAttribute('data-theme', theme),
            theme,
          )
          const tabs = layoutPage.locator('.ribbon-tabs')
          await expect(tabs).toHaveCSS('padding-left', '10px')
          await expect(tabs).toHaveCSS('padding-right', '10px')
          await expect(tabs).toHaveCSS('-webkit-app-region', 'no-drag')
          const file = await layoutPage
            .getByRole('button', { name: 'File', exact: true })
            .boundingBox()
          const save = await layoutPage.locator('.qa-btn').first().boundingBox()
          const row = await tabs.boundingBox()
          expect(file!.x - row!.x, `${platform} ${theme} File inset`).toBe(10)
          // Native File styling has a 6px margin plus the row's 2px gap.
          expect(save!.x - file!.x - file!.width, `${platform} ${theme} File/Save gap`).toBe(8)
          await layoutPage.screenshot({
            path: test.info().outputPath(`embedded-${platform.split(' ')[0]}-${theme}.png`),
          })
          await layoutPage.locator('.ProseMirror').first().click()
          await layoutPage.keyboard.insertText('Focus fixture')
          const fileButton = layoutPage.getByRole('button', { name: 'File', exact: true })
          await fileButton.click()
          const entries = layoutPage.getByRole('menuitem')
          await expect(entries.first()).toBeFocused()
          await expect(entries.first()).toHaveCSS('outline-style', 'none')
          await expect(entries.first()).toHaveCSS('background-color', 'rgba(0, 0, 0, 0)')
          await layoutPage.keyboard.press('ArrowDown')
          await expect(entries.nth(1)).toBeFocused()
          await expect(entries.nth(1)).toHaveCSS('outline-style', 'solid')
          const focusColor = await entries
            .nth(1)
            .evaluate((item) => getComputedStyle(item).outlineColor)
          await layoutPage.keyboard.press('Escape')
          await expect(fileButton).toBeFocused()
          await expect(fileButton).toHaveCSS('outline-style', 'solid')
          await expect(fileButton).toHaveCSS('outline-color', focusColor)
          await expect(fileButton).toHaveCSS('outline-offset', '2px')
          await layoutPage.keyboard.press('ArrowUp')
          await expect(entries.last()).toBeFocused()
          await layoutPage.keyboard.press('Home')
          await expect(entries.first()).toBeFocused()
          await layoutPage.keyboard.press('Tab')
          await expect(layoutPage.getByRole('menu')).toHaveCount(0)
          await expect(layoutPage.locator('.qa-btn').first()).toBeFocused()
          await layoutPage.keyboard.press('Shift+Tab')
          await expect(fileButton).toBeFocused()
          await expect(fileButton).toHaveCSS('outline-style', 'solid')
          await layoutPage.keyboard.press('Enter')
          await expect(entries.first()).toBeFocused()
          await expect(entries.first()).toHaveCSS('outline-style', 'solid')
          await layoutPage.keyboard.press('Escape')
        }
      } finally {
        await layoutPage.close()
      }
    }
    const page = await browser.newPage()
    page.on('pageerror', (error) => errors.push(error.message))
    await page.goto(`http://127.0.0.1:${address.port}/?empty`)
    await page.waitForFunction(() => (window as unknown as FixtureWindow).sdkReady)
    expect(
      await page.evaluate(() => (window as unknown as FixtureWindow).editor.documents.state()),
    ).toBeNull()
    expect(
      await page.evaluate(() => (window as unknown as FixtureWindow).editor.documents.new()),
    ).toBe(true)
    const body = page.locator('.ProseMirror').first()
    const autoSaveInput = page.locator('.autosave-toggle input')
    await expect(autoSaveInput).toBeEnabled()
    await expect(autoSaveInput).toBeChecked()
    await page.locator('.autosave-toggle').click()
    await expect(autoSaveInput).not.toBeChecked()
    await body.click()
    await page.keyboard.insertText('same same same [invalid')
    await page.evaluate(() => window.dispatchEvent(new Event('blur')))
    await page.waitForTimeout(350)
    expect(
      await page.evaluate(() => (window as unknown as FixtureWindow).fixture.saveKinds),
    ).toEqual([])
    expect(
      await page.evaluate(
        () => (window as unknown as { checkClose(): { autoSave: boolean } }).checkClose().autoSave,
      ),
    ).toBe(false)
    await page.locator('.autosave-toggle').click()
    await expect(autoSaveInput).toBeChecked()
    expect(
      await page.evaluate(
        () => (window as unknown as { checkClose(): { autoSave: boolean } }).checkClose().autoSave,
      ),
    ).toBe(true)
    const instructionInput = page.locator('.ai-ask-pop-input')
    const openSelectionPrompt = async () => {
      await body.click()
      await body.press('Control+Home')
      // Select the known fixture text without Electron's native word-navigation command.
      for (let index = 0; index < 'same'.length; index++) await body.press('Shift+ArrowRight')
      await expect.poll(() => page.evaluate(() => window.getSelection()?.toString())).toBe('same')
      await expect(page.locator('.ai-ask-trigger')).toBeVisible()
      await page.locator('.ai-ask-trigger').click()
      await expect(instructionInput).toBeFocused()
      await expect(page.locator('.ai-ask-pop-sub')).toHaveText('same')
      await expect(page.locator('.doc-inactive-selection')).toBeVisible()
    }
    for (const cancel of ['close', 'escape', 'outside']) {
      await openSelectionPrompt()
      await instructionInput.fill('Do not submit this')
      if (cancel === 'close') await page.locator('.ai-ask-pop-close').click()
      else if (cancel === 'escape') await page.keyboard.press('Escape')
      else await page.getByRole('button', { name: 'View', exact: true }).click()
      await expect(instructionInput).toHaveCount(0)
      if (cancel !== 'outside') await expect(body).toBeFocused()
      await expect(page.locator('.doc-inactive-selection')).toHaveCount(0)
      expect(
        await page.evaluate(() => (window as unknown as FixtureWindow).fixture.agentActions),
      ).toEqual([])
    }
    await page.getByRole('button', { name: 'Home', exact: true }).click()
    for (const confirm of ['button', 'enter']) {
      await openSelectionPrompt()
      await expect(page.locator('.ai-ask-confirm')).toBeDisabled()
      await instructionInput.fill('Polish this passage without changing its meaning')
      if (confirm === 'button') await page.locator('.ai-ask-confirm').click()
      else await page.keyboard.press('Enter')
      await expect(instructionInput).toHaveCount(0)
      await expect
        .poll(() =>
          page.evaluate(() => (window as unknown as FixtureWindow).fixture.agentActions.length),
        )
        .toBe(confirm === 'button' ? 1 : 2)
      const action = await page.evaluate(() =>
        (window as unknown as FixtureWindow).fixture.agentActions.at(-1)!,
      )
      expect(action).toMatchObject({
        kind: 'edit',
        instruction: 'Polish this passage without changing its meaning',
      })
      await body.click()
      await page.keyboard.press('Control+End')
      await page.keyboard.press('Control+Shift+ArrowLeft')
      await expect
        .poll(() =>
          page.evaluate(() => {
            const live = (window as unknown as FixtureWindow).editor.runtime.connect({
              id: 'await-live-selection',
            })
            try {
              return live.context().text
            } finally {
              live.close()
            }
          }),
        )
        .not.toBe(action.context)
      const binding = await page.evaluate(async (action) => {
        const { editor } = window as unknown as FixtureWindow
        const bound = await editor.commands.handle({
          requestId: crypto.randomUUID(),
          callerId: 'external-agent',
          bindingId: action.bindingId,
          command: 'read_document',
          payload: {},
        })
        const live = editor.runtime.connect({ id: 'live-selection' })
        try {
          return { bound, live: live.context() }
        } finally {
          live.close()
        }
      }, action)
      expect(binding.bound).toMatchObject({ documentId: action.documentId, text: action.context })
      expect((binding.bound as { text: string }).text).not.toBe(binding.live.text)
    }
    const assistant = page.getByRole('button', { name: 'Workspace assistant', exact: true })
    await expect(page.getByRole('button', { name: 'Genspark AI', exact: true })).toBeVisible()
    await page.evaluate((icon) => {
      const { editor } = window as unknown as FixtureWindow
      editor.ui.ribbon.set({
        documentToken: editor.documents.state()!.documentToken,
        assistantButton: { label: 'Workspace assistant', icon: `data:image/png;base64,${icon}` },
      })
    }, image)
    await expect(assistant).toBeVisible()
    await expect(assistant.locator('img')).toHaveAttribute('src', `data:image/png;base64,${image}`)
    await expect(assistant.locator('img')).toHaveCSS('width', '26px')
    await expect(assistant.locator('xpath=../..').locator('.ribbon-group-label')).toHaveText(
      'Workspace assistant',
    )
    await expect(page.getByRole('button', { name: 'Genspark AI', exact: true })).toHaveCount(0)
    await assistant.click()
    expect(
      await page.evaluate(() => (window as unknown as FixtureWindow).fixture.agentActions.at(-1)),
    ).toMatchObject({ kind: 'ask', instruction: '' })
    await page.evaluate(() => {
      const { editor } = window as unknown as FixtureWindow
      editor.ui.ribbon.set({
        documentToken: editor.documents.state()!.documentToken,
        assistantButton: { label: 'Workspace assistant', visible: false },
      })
    })
    await expect(assistant).toHaveCount(0)
    await expect(page.getByRole('button', { name: 'AI Summarize', exact: true })).toBeVisible()
    await page.evaluate(() => (window as unknown as FixtureWindow).editor.ui.ribbon.reset())
    await expect(page.getByRole('button', { name: 'Genspark AI', exact: true })).toBeVisible()
    const bodyInstance = await body.elementHandle()
    const beforeMenu = await page.evaluate(() => {
      const target = window as unknown as FixtureWindow
      const { editor } = target
      const state = editor.documents.state()!
      editor.ui.fileMenu.onAction((event) => target.fixture.menuActions.push(event))
      editor.ui.fileMenu.set({
        documentToken: state.documentToken,
        items: [
          { type: 'command', id: 'save', command: 'save' },
          {
            type: 'submenu',
            id: 'workflow',
            label: 'Workflow',
            items: [{ type: 'action', id: 'publish', actionId: 'host.publish', label: 'Publish' }],
          },
        ],
      })
      return { snapshot: editor.inspection.snapshot(), state }
    })
    await page.getByRole('button', { name: 'File', exact: true }).click()
    await expect(page.getByRole('menu', { name: 'File', exact: true })).toBeVisible()
    await page.evaluate((icon) => {
      const { editor, fixture } = window as unknown as FixtureWindow
      const configuration: DocsRibbonConfiguration = {
        documentToken: editor.documents.state()!.documentToken,
        tabs: [
          {
            id: 'extension.tools',
            label: 'Extension tools',
            groups: [
              {
                id: 'extension-group',
                label: 'Extension',
                items: [
                  {
                    type: 'button',
                    id: 'inspect',
                    label: 'Inspect',
                    actionId: 'extension.inspect',
                    pressed: false,
                  },
                ],
              },
            ],
          },
        ],
        groups: [
          {
            tabId: 'home',
            id: 'extension-home',
            label: 'Home tools',
            items: [
              {
                type: 'button',
                id: 'home-action',
                label: 'Home action',
                actionId: 'extension.home',
              },
              {
                type: 'button',
                id: 'home-parity-image',
                label: 'Convert',
                actionId: 'extension.parity',
                icon: `data:image/png;base64,${icon}`,
              },
              {
                type: 'button',
                id: 'home-parity-fallback',
                label: 'Convert',
                actionId: 'extension.parity',
              },
              {
                type: 'menu',
                id: 'home-parity-menu',
                label: 'Convert',
                items: [
                  {
                    type: 'action',
                    id: 'home-parity-action',
                    label: 'Run',
                    actionId: 'extension.parity',
                  },
                ],
              },
            ],
          },
          {
            tabId: 'view',
            id: 'extension-view',
            label: 'View tools',
            items: [
              {
                type: 'menu',
                id: 'view-menu',
                label: 'View actions',
                items: [
                  {
                    type: 'action',
                    id: 'view-action',
                    label: 'View action',
                    actionId: 'extension.view',
                  },
                  {
                    type: 'submenu',
                    id: 'view-nested',
                    label: 'Nested actions',
                    items: [
                      {
                        type: 'action',
                        id: 'nested-action',
                        label: 'Nested action',
                        actionId: 'extension.nested',
                      },
                    ],
                  },
                ],
              },
              {
                type: 'button',
                id: 'native-parity',
                label: 'Read Mode',
                actionId: 'extension.parity',
                icon: `data:image/png;base64,${icon}`,
              },
            ],
          },
        ],
        quickAccess: [
          {
            type: 'button',
            id: 'quick-action',
            label: 'Quick action',
            actionId: 'extension.quick',
            pressed: true,
            icon: `data:image/png;base64,${icon}`,
          },
        ],
      }
      const setPanelOpen = (open: boolean) => {
        const button = configuration.tabs![0].groups[0].items[0]
        if (button.type === 'button') button.pressed = open
        editor.ui.ribbon.set(configuration)
      }
      ;(window as unknown as FixtureWindow).setExtensionPanelOpen = setPanelOpen
      editor.ui.ribbon.onAction((event) => {
        fixture.ribbonActions.push(event)
        if (event.actionId === 'extension.inspect') setPanelOpen(true)
      })
      editor.ui.ribbon.set(configuration)
    }, image)
    // Updating external controls must leave an unrelated native menu open.
    await expect(page.getByRole('menu', { name: 'File', exact: true })).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(page.getByRole('button', { name: 'Home', exact: true })).toBeVisible()
    const measureHomeParity = () =>
      page.evaluate(() => {
        const measure = (button: Element) => {
          const style = getComputedStyle(button)
          const box = button.getBoundingClientRect()
          const glyph = button.querySelector('svg,img')!.getBoundingClientRect()
          return {
            height: box.height,
            top: box.top,
            glyphWidth: glyph.width,
            glyphHeight: glyph.height,
            font: style.font,
            padding: style.padding,
            gap: style.gap,
            plateHeight: button.querySelector('.rb-big-icon')!.getBoundingClientRect().height,
          }
        }
        const native = Array.from(
          document.querySelectorAll('.ribbon-group:not([data-ribbon-extension]) .rb-big'),
        ).find((button) => button.textContent?.trim() === 'Convert')!
        return {
          native: measure(native),
          styles: measure(document.querySelector('.rb-styles-gallery')!),
          extensions: Array.from(
            document.querySelectorAll(
              '[data-ribbon-extension="extension-home"] button[aria-label="Convert"]',
            ),
          ).map(measure),
        }
      })
    for (const theme of ['light', 'dark']) {
      await page.evaluate(
        (theme) => document.documentElement.setAttribute('data-theme', theme),
        theme,
      )
      for (const compact of [false, true]) {
        if (compact) await page.keyboard.press('Control+Alt+k')
        const parity = await measureHomeParity()
        expect(parity.native.glyphWidth).toBe(22)
        expect(parity.styles).toEqual(parity.native)
        expect(parity.extensions).toHaveLength(3)
        for (const extension of parity.extensions) expect(extension).toEqual(parity.native)
      }
      await page.keyboard.press('Control+Alt+k')
    }
    await page.evaluate(() => document.documentElement.setAttribute('data-theme', 'light'))
    await page.getByRole('button', { name: 'Home action', exact: true }).click()
    await page.getByRole('button', { name: 'Extension tools', exact: true }).click()
    const panelButton = page.getByRole('button', { name: 'Inspect', exact: true })
    await expect(panelButton).toHaveAttribute('aria-pressed', 'false')
    const releasedColor = await panelButton
      .locator('.rb-big-icon')
      .evaluate((plate) => getComputedStyle(plate).backgroundColor)
    await panelButton.click()
    await expect(panelButton).toHaveAttribute('aria-pressed', 'true')
    await expect(panelButton).toHaveClass(/active/)
    expect(
      await panelButton
        .locator('.rb-big-icon')
        .evaluate((plate) => getComputedStyle(plate).backgroundColor),
    ).not.toBe(releasedColor)
    await page.evaluate(() => (window as unknown as FixtureWindow).setExtensionPanelOpen(false))
    await expect(panelButton).toHaveAttribute('aria-pressed', 'false')
    await expect(panelButton).not.toHaveClass(/active/)
    const quickButton = page.getByRole('button', { name: 'Quick action', exact: true })
    const pressedQuickColor = await quickButton.evaluate(
      (button) => getComputedStyle(button).backgroundColor,
    )
    await quickButton.click()
    expect(await quickButton.evaluate((button) => getComputedStyle(button).backgroundColor)).toBe(
      pressedQuickColor,
    )
    await expect(page.getByRole('button', { name: 'Quick action', exact: true })).toHaveAttribute(
      'aria-pressed',
      'true',
    )
    await page.getByRole('button', { name: 'View', exact: true }).click()
    const measureNativeParity = () =>
      page.evaluate(() => {
        const measure = (button: HTMLButtonElement) => {
          const style = getComputedStyle(button)
          const box = button.getBoundingClientRect()
          const plate = button.querySelector('.rb-big-icon')!.getBoundingClientRect()
          const glyph = button.querySelector('svg,img')!.getBoundingClientRect()
          return {
            width: box.width,
            height: box.height,
            top: box.top,
            font: style.font,
            lineHeight: style.lineHeight,
            padding: style.padding,
            gap: style.gap,
            plateHeight: plate.height,
            glyphWidth: glyph.width,
            glyphHeight: glyph.height,
          }
        }
        return {
          native: measure(
            Array.from(
              document.querySelectorAll<HTMLButtonElement>(
                '.ribbon-body > .ribbon-group:not([data-ribbon-extension]) .rb-big',
              ),
            ).find((button) => button.textContent?.trim() === 'Read Mode')!,
          ),
          extension: measure(
            document.querySelector<HTMLButtonElement>(
              '[data-ribbon-extension="extension-view"] button[aria-label="Read Mode"]',
            )!,
          ),
        }
      })
    const fullParity = await measureNativeParity()
    expect(fullParity.extension).toEqual(fullParity.native)
    const extensionTrigger = page.getByRole('button', { name: 'View actions', exact: true })
    await extensionTrigger.click()
    const extensionPopup = page.getByRole('menu', { name: 'View actions', exact: true })
    await expect(extensionPopup).toBeVisible()
    const menuColors: string[] = []
    for (const theme of ['light', 'dark']) {
      await page.evaluate(
        (theme) => document.documentElement.setAttribute('data-theme', theme),
        theme,
      )
      menuColors.push(
        await extensionPopup.evaluate((menu) => getComputedStyle(menu).backgroundColor),
      )
      await page.screenshot({ path: test.info().outputPath(`ribbon-${theme}.png`) })
    }
    expect(menuColors[0]).not.toBe(menuColors[1])
    await page.evaluate(() => document.documentElement.setAttribute('data-theme', 'light'))
    await page.keyboard.press('Control+Alt+k')
    await expect(page.locator('.ribbon')).toHaveClass(/ribbon-compact/)
    const compactParity = await measureNativeParity()
    expect(compactParity.extension).toEqual(compactParity.native)
    await expect(extensionTrigger).toBeVisible()
    await expect(extensionPopup).toBeVisible()
    await page.keyboard.press('Control+Alt+k')
    const band = await page.locator('.ribbon-body').boundingBox()
    const popup = await extensionPopup.boundingBox()
    expect(popup!.y).toBeGreaterThanOrEqual(band!.y)
    await page.getByRole('menuitem', { name: 'View action', exact: true }).click()
    await expect(extensionPopup).toBeHidden()
    await page.mouse.move(0, 0)
    await extensionTrigger.focus()
    await page.keyboard.press('ArrowDown')
    await expect(page.getByRole('menuitem', { name: 'View action', exact: true })).toBeFocused()
    await page.keyboard.press('ArrowDown')
    await expect(page.getByRole('menuitem', { name: 'Nested actions', exact: true })).toBeFocused()
    await page.keyboard.press('ArrowRight')
    await expect(page.getByRole('menuitem', { name: 'Nested action', exact: true })).toBeFocused()
    await page.keyboard.press('Enter')
    await expect(extensionTrigger).toBeFocused()
    expect(
      await page.evaluate(() =>
        (window as unknown as FixtureWindow).fixture.ribbonActions.map((event) => [
          event.actionId,
          event.tabId,
          event.location,
        ]),
      ),
    ).toEqual([
      ['extension.home', 'home', 'tab'],
      ['extension.inspect', 'extension.tools', 'tab'],
      ['extension.quick', null, 'quickAccess'],
      ['extension.view', 'view', 'tab'],
      ['extension.nested', 'view', 'tab'],
    ])
    expect(
      await page.evaluate(() => {
        const { editor } = window as unknown as FixtureWindow
        return { snapshot: editor.inspection.snapshot(), state: editor.documents.state() }
      }),
    ).toEqual(beforeMenu)
    // Removing the selected custom tab returns to Home without remounting the document.
    await page.getByRole('button', { name: 'Extension tools', exact: true }).click()
    await page.evaluate(() => (window as unknown as FixtureWindow).editor.ui.ribbon.reset())
    await expect(page.getByRole('button', { name: 'Extension tools', exact: true })).toHaveCount(0)
    await expect(page.getByRole('button', { name: 'Home', exact: true })).toHaveClass(/active/)
    await expect(page.getByRole('button', { name: 'Quick action', exact: true })).toHaveCount(0)

    const fileButton = page.getByRole('button', { name: 'File', exact: true })
    const fileBox = await fileButton.boundingBox()
    const saveBox = await page.locator('.qa-btn').first().boundingBox()
    expect(fileBox!.x + fileBox!.width).toBeLessThanOrEqual(saveBox!.x + 1)
    await fileButton.focus()
    await page.keyboard.press('Enter')
    await expect(page.getByRole('menuitem', { name: /^Save/ })).toBeFocused()
    await page.keyboard.press('ArrowDown')
    await page.keyboard.press('ArrowRight')
    await expect(page.getByRole('menuitem', { name: 'Publish', exact: true })).toBeFocused()
    await expect(page.getByRole('menuitem', { name: 'Publish', exact: true })).toHaveCSS(
      'outline-style',
      'solid',
    )
    await page.keyboard.press('Escape')
    await expect(page.getByRole('menuitem', { name: 'Workflow' })).toBeFocused()
    await page.keyboard.press('Escape')
    await expect(fileButton).toBeFocused()
    await fileButton.click()
    await page.getByRole('menuitem', { name: 'Workflow' }).click()
    await page.getByRole('menuitem', { name: 'Publish', exact: true }).click()
    expect(
      await page.evaluate(() => (window as unknown as FixtureWindow).fixture.menuActions),
    ).toEqual([
      expect.objectContaining({
        actionId: 'host.publish',
        documentToken: beforeMenu.state.documentToken,
      }),
    ])
    await page.evaluate(() => {
      const { editor } = window as unknown as FixtureWindow
      editor.ui.fileMenu.set({
        documentToken: editor.documents.state()!.documentToken,
        items: [
          { type: 'command', id: 'save', command: 'save', visible: false },
          {
            type: 'action',
            id: 'publish',
            actionId: 'host.publish',
            label: 'Published',
            enabled: false,
          },
        ],
      })
    })
    await fileButton.click()
    await expect(page.getByRole('menuitem', { name: 'Published' })).toBeDisabled()
    await expect(page.getByRole('menuitem', { name: /^Save/ })).toHaveCount(0)
    expect(
      await page.evaluate(() => {
        const { editor } = window as unknown as FixtureWindow
        return { snapshot: editor.inspection.snapshot(), state: editor.documents.state() }
      }),
    ).toEqual(beforeMenu)
    expect(await body.evaluate((element, original) => element === original, bodyInstance)).toBe(
      true,
    )
    await page.evaluate(() => window.dispatchEvent(new Event('blur')))
    await expect
      .poll(() =>
        page.evaluate(() => (window as unknown as FixtureWindow).editor.documents.state()?.dirty),
      )
      .toBe(false)
    expect(
      await page.evaluate(() => (window as unknown as FixtureWindow).fixture.saveKinds),
    ).toContain(true)
    await body.click()
    await page.keyboard.insertText(' ')
    await page.keyboard.press('Control+s')
    await expect
      .poll(() =>
        page.evaluate(() => (window as unknown as FixtureWindow).editor.documents.state()?.dirty),
      )
      .toBe(false)
    await page.evaluate(() => (window as unknown as FixtureWindow).editor.ui.fileMenu.reset())
    const results = await page.evaluate(async () => {
      const editor = (window as unknown as FixtureWindow).editor
      const found = editor.inspection.find('(same)', { mode: 'regex' })
      editor.inspection.locate(found.matches[1])
      editor.annotations.setRules('display', [
        { id: 'all', pattern: 'same', style: { background: 'yellow' } },
      ])
      await editor.annotations.refresh()
      const diagnostics = editor.annotations.diagnostics()
      const count = editor.inspection.replace(found, '$1!')
      await editor.documents.save()
      return {
        count,
        diagnostics,
        text: editor.inspection.snapshot().text,
        state: editor.documents.state(),
        bytes: (window as unknown as FixtureWindow).fixture.saved,
        fallback: (window as unknown as FixtureWindow).fixture.fallback,
      }
    })
    expect(results.count).toBe(3)
    expect(results.diagnostics).toHaveLength(1)
    expect(results.text).toContain('same! same! same!')
    expect(results.state!.path).toBeNull()
    expect(results.fallback).toEqual([])
    const archive = await JSZip.loadAsync(Uint8Array.from(results.bytes))
    expect(await archive.file('word/document.xml')!.async('string')).not.toContain(
      'sdk-range-decoration',
    )
    if (sdkHostIsElectron) {
      await page.evaluate(async () => {
        await document.fonts.ready
        await (window as unknown as FixtureWindow).editor.annotations.refresh()
      })
      await expect(page.locator('.sdk-range-decoration').first()).toBeVisible()
      await expect(page.locator('.app-toast')).toHaveCount(0)
      await body.evaluate((element) => (element as HTMLElement).blur())
      const annotated = await browser.printToPDF!(page)
      await expect(page.locator('.sdk-range-decoration').first()).toBeVisible()
      await page.evaluate(() => (window as unknown as FixtureWindow).editor.annotations.reset())
      await expect(page.locator('.sdk-range-decoration')).toHaveCount(0)
      await body.evaluate((element) => (element as HTMLElement).blur())
      const clean = await browser.printToPDF!(page)
      await writeFile(test.info().outputPath('annotated.pdf'), annotated)
      await writeFile(test.info().outputPath('clean.pdf'), clean)
      const render = await createPrintRenderer()
      const decorated = render(annotated)
      expect(decorated.length).toBeGreaterThan(0)
      await writeFile(test.info().outputPath('annotated-print-page-1.png'), decorated[0].png)
      expect(
        decorated
          .map((page) => page.text)
          .join('')
          .replace(/\s+/g, ''),
      ).toContain('same!same!same!')
      const plain = render(clean)
      await writeFile(test.info().outputPath('clean-print-page-1.png'), plain[0].png)
      expect(decorated.map((page) => page.text)).toEqual(plain.map((page) => page.text))
      expect(decorated.length).toBe(plain.length)
      for (const [index, page] of decorated.entries())
        expect(
          Buffer.from(page.png).equals(Buffer.from(plain[index].png)),
          `print page ${index}`,
        ).toBe(true)
    }
    await body.click()
    await page.keyboard.press('Control+Home')
    await page.keyboard.press('Control+Shift+ArrowRight')
    await page.locator('.ai-ask-trigger').click()
    await instructionInput.fill('An unfinished request belongs to this document')
    await page.evaluate(async () => {
      const target = window as unknown as FixtureWindow
      target.fixture.readonly = true
      await target.editor.documents.open()
    })
    await expect(instructionInput).toHaveCount(0)
    expect(
      await page.evaluate(() => (window as unknown as FixtureWindow).fixture.agentActions.length),
    ).toBe(3)
    expect(
      await page.evaluate(async () => {
        const { editor, fixture } = window as unknown as FixtureWindow
        try {
          await editor.commands.handle({
            requestId: crypto.randomUUID(),
            callerId: 'external-agent',
            bindingId: fixture.agentActions[0].bindingId,
            command: 'read_document',
            payload: {},
          })
          return false
        } catch {
          return true
        }
      }),
    ).toBe(true)
    expect(
      await page.evaluate((token) => {
        const { editor } = window as unknown as FixtureWindow
        try {
          editor.ui.fileMenu.set({ documentToken: token, items: [] })
          return false
        } catch {
          return true
        }
      }, beforeMenu.state.documentToken),
    ).toBe(true)
    await expect(body).toHaveAttribute('contenteditable', 'false')
    await expect(autoSaveInput).toBeDisabled()
    await expect(autoSaveInput).not.toBeChecked()
    expect(
      await page.evaluate(
        () => (window as unknown as { checkClose(): { autoSave: boolean } }).checkClose().autoSave,
      ),
    ).toBe(false)
    expect(
      await page.evaluate(
        () => (window as unknown as FixtureWindow).editor.documents.state()?.dirty,
      ),
    ).toBe(false)
    expect(
      await page.evaluate(() => (window as unknown as FixtureWindow).fixture.saveKinds.at(-1)),
    ).toBe(false)
    await page.evaluate(() =>
      (window as unknown as FixtureWindow).editor.annotations.setRules('readonly', [
        { id: 'read', pattern: 'same', style: { underline: true } },
      ]),
    )
    await expect(page.locator('.sdk-range-decoration').first()).toBeVisible()
    await page.evaluate(async () => {
      const target = window as unknown as FixtureWindow
      target.editor.documents.subscribe((state) =>
        target.fixture.documentStates.push(!!state?.readOnly),
      )
      target.fixture.readonly = false
      await target.editor.documents.open()
    })
    await expect(body).toHaveAttribute('contenteditable', 'true')
    expect(
      await page.evaluate(() => (window as unknown as FixtureWindow).fixture.documentStates.at(-1)),
    ).toBe(false)
    expect(
      await page.evaluate(() =>
        (window as unknown as FixtureWindow).fixture.documentStates.slice(1),
      ),
    ).not.toContain(true)
    expect(errors).toEqual([])
    expect(
      await page.evaluate(() =>
        (
          window as unknown as { remountSdk(): Promise<{ refused: boolean; calls: number }> }
        ).remountSdk(),
      ),
    ).toEqual({ refused: true, calls: 0, oldOpenResult: false })
  } finally {
    await browser?.close()
    await new Promise<void>((done, reject) =>
      server.close((error) => (error ? reject(error) : done())),
    )
  }
})

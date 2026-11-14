import { test, expect, _electron as electron } from '@playwright/test'
import { createServer } from 'node:http'
import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { extname, join, resolve, sep } from 'node:path'
import JSZip from 'jszip'
import { createRequire } from 'node:module'
import { randomUUID } from 'node:crypto'
import type { DocsEditorHandle } from '../apps/docs/src/renderer/embedded'
import type { DocsQueryResult } from '../apps/docs/src/shared/document-query'
import type { DocsEditorAction } from '../apps/docs/src/shared/editor-action'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js'
import { createDocsHostBridge, type DocsHostBridge } from '../apps/docs/src/services/host-bridge'
import { DocsOperationError, isDocsOperationResult } from '../apps/docs/src/shared/operation-result'
import type { DesktopApi, McpCommandResult, McpSaveResult } from '../apps/docs/src/shared/ipc'
import { runRealCodexDocs } from './real-codex'
import { runRealDshDocs } from './real-dsh'
import { LIVE_DOCS_TOOL_NAMES } from '../apps/shell/tests/mcp/docs-tool-names'

declare global {
  interface Window {
    docsTest: {
      editor: DocsEditorHandle
      actions: DocsEditorAction[]
      ready: boolean
      services: DesktopApi
    }
    docsPortableSave(
      path: string,
      bytes: number[],
      overwrite: boolean,
      sourcePath?: string | null,
    ): Promise<McpSaveResult>
  }
}

const loadServices = createRequire(__filename)
const sdkRoot = resolve(__dirname, '../apps/docs/out/embedded')

test('the embedded public editor edits, undoes, saves and reopens Word with its native Agent disabled', async () => {
  test.setTimeout(180_000)
  expect(existsSync(join(sdkRoot, 'editor.js')), 'Build the embeddable Docs editor first').toBe(
    true,
  )
  const scratch = await mkdtemp(join(tmpdir(), 'foraoffice-docs-embedded-'))
  const output = join(scratch, 'shared-runtime.docx')
  const documents = join(scratch, 'documents')
  const userData = join(scratch, 'electron-state')
  await mkdir(documents)
  await mkdir(userData)
  const backend: typeof import('../apps/docs/src/services') = loadServices(
    resolve(__dirname, '../apps/docs/out/services/services.cjs'),
  )
  const files = await backend.createDocsFileSession({
    stateDirectory: join(scratch, 'node-state'),
    readRoots: [documents],
    writeRoots: [documents],
  })
  const fileServices = backend.createDocsFileServices({
    files,
    defaultSaveDirectory: documents,
    dialogs: {
      pickOpen: async () => null,
      pickSave: async () => null,
      chooseRecovery: async () => 'discard',
      confirmExternalOverwrite: async () => false,
    },
    recentFiles: { list: async () => [], add: async () => {} },
  })
  await writeFile(
    join(userData, 'app-settings.json'),
    JSON.stringify({ defaultSaveDir: documents }),
  )
  const server = createServer(async (request, response) => {
    try {
      const pathname = new URL(request.url ?? '/', 'http://localhost').pathname
      if (pathname === '/') {
        response.setHeader('content-type', 'text/html')
        response.end(`<!doctype html><html><head><link rel="stylesheet" href="/sdk/editor.css"></head>
          <body><div id="root"></div><script type="module">
          import { mountDocsEditor, createExternalDocsHost } from '/sdk/editor.js';
          localStorage.setItem('aidocs.autoSave', '0');
          window.docsTest = { actions: [], ready: false };
          window.docsTest.services = Object.create(window.desktop);
          const host = createExternalDocsHost('electron', window.docsTest.services, { submit: async action => { window.docsTest.actions.push(action); } });
          window.docsTest.editor = await mountDocsEditor(document.getElementById('root'), { host });
          await window.docsTest.editor.ready;
          window.docsTest.ready = true;
          </script></body></html>`)
        return
      }
      if (!pathname.startsWith('/sdk/')) {
        response.writeHead(404).end()
        return
      }
      const file = resolve(sdkRoot, decodeURIComponent(pathname.slice(5)))
      if (!file.startsWith(sdkRoot + sep)) {
        response.writeHead(403).end()
        return
      }
      response.setHeader(
        'content-type',
        (
          {
            '.js': 'text/javascript',
            '.css': 'text/css',
            '.woff2': 'font/woff2',
            '.wasm': 'application/wasm',
          } as Record<string, string>
        )[extname(file)] ?? 'application/octet-stream',
      )
      response.end(await readFile(file))
    } catch {
      response.writeHead(404).end()
    }
  })
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('No fixture server port')
  const app = await electron.launch({
    args: [resolve(__dirname, '../apps/docs')],
    env: {
      ...process.env,
      AI_OFFICE_USER_DATA: userData,
      GENOFFICE_LANG: 'en',
      ELECTRON_RENDERER_URL: `http://127.0.0.1:${address.port}`,
    },
  })
  let documentBridge: DocsHostBridge | undefined
  let mcp: Client | undefined
  const requests: string[] = []
  const successfulTools = new Set<string>()
  try {
    await app.evaluate(({ dialog }, path) => {
      dialog.showSaveDialog = async () => ({ canceled: false, filePath: path })
      dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [path] })
    }, output)
    const page = await app.firstWindow()
    await page.exposeFunction(
      'docsPortableSave',
      (path: string, bytes: number[], overwrite: boolean, sourcePath?: string | null) =>
        fileServices.saveDocxTo(path, Uint8Array.from(bytes).buffer, overwrite, sourcePath),
    )
    await page.waitForFunction(() => window.docsTest?.ready)
    await expect(page.locator('.ai-dock')).toHaveCount(0)
    const body = page.locator('.ProseMirror').first()
    await body.click()
    await page.keyboard.insertText('Manual Word text')
    await page.keyboard.press('Control+Home')
    for (let i = 0; i < 6; i++) await page.keyboard.press('Shift+ArrowRight')
    await page.locator('.ai-ask-trigger').click()
    await page.locator('.ai-ask-pop-input').fill('Edit the selected text')
    await page.locator('.ai-ask-confirm').click()
    await expect.poll(() => page.evaluate(() => window.docsTest.actions.length)).toBe(1)
    const bindingId = await page.evaluate(() => window.docsTest.actions[0].bindingId)
    const createBridge = (bindingId: string) =>
      createDocsHostBridge({
        bindingId,
        releaseScope: (binding) => {
          void page
            .evaluate((id) => window.docsTest.editor.commands.releaseScope(id), binding)
            .catch(() => {})
        },
        endpoint: {
          request: async (message, signal) => {
            requests.push(
              message.command === 'execute_tool'
                ? String((message.payload as { name?: string }).name)
                : message.command,
            )
            const cancel = () => {
              void page
                .evaluate(
                  (request) =>
                    window.docsTest.editor.commands.handle({
                      ...request,
                      command: 'cancel_operation',
                      payload: { requestId: request.requestId },
                    }),
                  message,
                )
                .catch(() => {})
            }
            signal.addEventListener('abort', cancel, { once: true })
            try {
              const reply: McpCommandResult = await page.evaluate(async (request) => {
                try {
                  return {
                    requestId: request.requestId,
                    ok: true,
                    result: await window.docsTest.editor.commands.handle(request),
                  }
                } catch (error) {
                  return {
                    requestId: request.requestId,
                    ok: false,
                    error: error instanceof Error ? error.message : String(error),
                    result: error instanceof Error && 'result' in error ? error.result : undefined,
                  }
                }
              }, message)
              if (!reply.ok) {
                if (isDocsOperationResult(reply.result))
                  throw new DocsOperationError(reply.error ?? reply.result.output, reply.result)
                throw new Error(reply.error)
              }
              return reply.result
            } finally {
              signal.removeEventListener('abort', cancel)
            }
          },
          closeCaller: (callerId) => {
            void page
              .evaluate((id) => window.docsTest.editor.commands.disconnect(id), callerId)
              .catch(() => {})
          },
        },
      })
    documentBridge = await createBridge(bindingId)
    const grantForHarness = async () => {
      // An independent conversation receives its own UI-selected scope and caller lifetime.
      const count = await page.evaluate(() => window.docsTest.actions.length)
      await body.click()
      await page.keyboard.press('Control+Home')
      for (let index = 0; index < 'Agent'.length; index++)
        await page.keyboard.press('Shift+ArrowRight')
      await page.locator('.ai-ask-trigger').click()
      await page.locator('.ai-ask-pop-input').fill('Edit the selected text')
      await page.locator('.ai-ask-confirm').click()
      await expect.poll(() => page.evaluate(() => window.docsTest.actions.length)).toBe(count + 1)
      return createBridge(await page.evaluate(() => window.docsTest.actions.at(-1)!.bindingId))
    }
    mcp = new Client({ name: 'packaged-stdio-acceptance', version: '1' }, { capabilities: {} })
    await mcp.connect(
      new StdioClientTransport({
        command: process.execPath,
        args: [resolve(__dirname, '../apps/docs/out/services/mcp-stdio.cjs')],
        env: {
          ...Object.fromEntries(
            Object.entries(process.env).filter(
              (entry): entry is [string, string] => entry[1] !== undefined,
            ),
          ),
          FORAOFFICE_DOCS_ENDPOINT: documentBridge.url,
          FORAOFFICE_DOCS_TOKEN: documentBridge.token,
        },
      }),
    )
    await page.keyboard.press('Control+End')
    const tool = async (name: string, input: Record<string, unknown>) => {
      const result = await mcp!.callTool({ name, arguments: input })
      if (result.isError) throw new Error(JSON.stringify(result.content))
      successfulTools.add(name)
      return result
    }
    const toolOutput = async (name: string, input: Record<string, unknown> = {}) => {
      const result = (await tool(name, input)).structuredContent
      if (!isDocsOperationResult(result))
        throw new Error(`${name} returned invalid operation metadata`)
      return result.output
    }
    expect(JSON.stringify((await tool('read_document', {})).content)).toContain('Manual Word text')
    const queried = JSON.parse(
      await toolOutput('query_document', {
        domain: 'body',
        containsText: 'Manual Word text',
        limit: 1,
      }),
    ) as DocsQueryResult
    expect(queried.page.complete).toBe(true)
    expect(queried.items).toHaveLength(1)
    expect(queried.items[0].html).toContain('Manual Word text')
    expect(queried.items[0].supportedOperations).toContain('findReplace')
    await tool('get_editor_capabilities', {})
    await tool('replace_selection', { html: 'Agent' })
    await expect(body).toHaveText('Agent Word text')
    await tool('document_history', { action: 'undo' })
    await expect(body).toHaveText('Manual Word text')
    await tool('document_history', { action: 'redo' })
    await expect(body).toHaveText('Agent Word text')
    const draft = await tool('start_document_draft', { afterBlockIndex: 0 })
    const draftId = JSON.parse((draft.structuredContent as { output: string }).output)
      .draftId as string
    await tool('update_document_draft', {
      draftId,
      sequence: 1,
      html: '<p>Visible streamed draft.</p>',
    })
    expect(await toolOutput('inspect_document_draft', { draftId })).toContain(draftId)
    await expect(body).toContainText('Visible streamed draft.')
    const previewSave = await mcp.callTool({ name: 'save_document', arguments: { path: output } })
    expect(previewSave.isError).toBe(true)
    expect(previewSave.structuredContent).toMatchObject({
      committed: false,
      errorCode: 'invalid_input',
    })
    expect(existsSync(output)).toBe(false)
    await tool('stop_document_draft', { draftId })
    const lateDraft = await mcp.callTool({
      name: 'update_document_draft',
      arguments: { draftId, sequence: 2, html: '<p>Late chunk.</p>' },
    })
    expect(lateDraft.isError).toBe(true)
    await tool('finish_document_draft', { draftId, decision: 'retain_partial' })
    await expect(body).toContainText('Visible streamed draft.')
    await tool('document_history', { action: 'undo' })
    await expect(body).toHaveText('Agent Word text')
    await page.evaluate(() => {
      // Electron freezes preload properties, including those inherited by this host.
      Object.defineProperty(window.docsTest.services, 'saveDocxTo', {
        configurable: true,
        value: (path: string, bytes: ArrayBuffer, overwrite: boolean, sourcePath?: string | null) =>
          window.docsPortableSave(path, Array.from(new Uint8Array(bytes)), overwrite, sourcePath),
      })
    })
    const denied = join(scratch, 'outside-grant.docx')
    const refusal = await mcp.callTool({ name: 'save_document', arguments: { path: denied } })
    expect(refusal.isError).toBe(true)
    expect(refusal.structuredContent).toMatchObject({
      committed: false,
      pending: true,
      errorCode: 'permission_denied',
    })
    expect(existsSync(denied)).toBe(false)
    const firstSave = join(documents, 'node-first-save.docx')
    const saved = await tool('save_document', { path: firstSave })
    expect(saved.structuredContent).toMatchObject({ committed: true, pending: false })
    expect(
      await (
        await JSZip.loadAsync(await readFile(firstSave))
      )
        .file('word/document.xml')!
        .async('string'),
    ).toContain('Agent')
    if (process.env.FORAOFFICE_CODEX_NODE_ENTRY || process.env.FORAOFFICE_CODEX_WSL_BINARY) {
      const requestStart = requests.length
      const codexOutput = join(documents, 'codex-saved.docx')
      const codexDirectory = join(scratch, 'codex')
      await mkdir(codexDirectory)
      const windowsSandbox = process.env.FORAOFFICE_CODEX_WINDOWS_SANDBOX
      if (windowsSandbox !== undefined && windowsSandbox !== 'elevated' && windowsSandbox !== 'mxc')
        throw new Error('Unsupported Codex acceptance sandbox selection')
      const wslBinary = process.env.FORAOFFICE_CODEX_WSL_BINARY
      const wslHome = process.env.FORAOFFICE_CODEX_WSL_HOME
      if (wslBinary && !wslHome)
        throw new Error('WSL acceptance needs its isolated authenticated Codex home')
      const bridge = await grantForHarness()
      try {
        await runRealCodexDocs({
          entry: process.env.FORAOFFICE_CODEX_NODE_ENTRY,
          cwd: codexDirectory,
          bridge,
          windowsSandbox,
          wsl: wslBinary
            ? {
                binary: wslBinary,
                home: wslHome!,
                distribution: process.env.FORAOFFICE_CODEX_WSL_DISTRO ?? 'Ubuntu',
              }
            : undefined,
          prompt: `Use the connected ForaOffice MCP tools to read the current live Word document, change only the captured selected word from Agent to Codex, and save the live document to ${JSON.stringify(codexOutput)}. Keep all other text. Use the editing tools, with no filesystem or shell commands. Report completion only after the tools confirm the save.`,
        })
      } finally {
        await bridge.close()
      }
      expect(requests.slice(requestStart)).toContain('save_document')
      await expect(body).toHaveText('Codex Word text')
      expect(
        await (
          await JSZip.loadAsync(await readFile(codexOutput))
        )
          .file('word/document.xml')!
          .async('string'),
      ).toContain('Codex')
      await tool('read_document', {})
      await tool('replace_selection', { html: 'Agent' })
      await expect(body).toHaveText('Agent Word text')
    }
    if (process.env.FORAOFFICE_DSH_NODE) {
      const packages = process.env.FORAOFFICE_DSH_PACKAGES
      const credentials = process.env.FORAOFFICE_DSH_CREDENTIALS
      if (!packages || !credentials)
        throw new Error(
          'DSH acceptance requires the installed packages and Harness credential-store paths',
        )
      const requestStart = requests.length
      const dshOutput = join(documents, 'dsh-saved.docx')
      const dshDirectory = join(scratch, 'dsh')
      await mkdir(dshDirectory)
      const bridge = await grantForHarness()
      try {
        await runRealDshDocs({
          node: process.env.FORAOFFICE_DSH_NODE,
          packages,
          credentials,
          cwd: dshDirectory,
          bridge,
          prompt: `Use the connected ForaOffice MCP tools to read the current live Word document, change only the captured selected word from Agent to DSH, and save the live document to ${JSON.stringify(dshOutput)}. Keep all other text. Use the editing tools and report completion only after the tools confirm the save.`,
        })
      } finally {
        await bridge.close()
      }
      expect(requests.slice(requestStart)).toContain('save_document')
      await expect(body).toHaveText('DSH Word text')
      expect(
        await (
          await JSZip.loadAsync(await readFile(dshOutput))
        )
          .file('word/document.xml')!
          .async('string'),
      ).toContain('DSH')
      await tool('read_document', {})
      await tool('replace_selection', { html: 'Agent' })
      await expect(body).toHaveText('Agent Word text')
    }
    // Continue exercising the native adapter with the same editor and runtime.
    await page.evaluate(() => {
      delete (window.docsTest.services as Partial<DesktopApi>).saveDocxTo
    })
    await page.keyboard.press('Control+Shift+S')
    await expect.poll(() => existsSync(output)).toBe(true)
    await expect(page.getByText('Saved', { exact: true })).toBeVisible()
    const archive = await JSZip.loadAsync(await readFile(output))
    const xml = await archive.file('word/document.xml')!.async('string')
    expect(xml).toContain('Agent')
    expect(xml).toContain('Word text')
    const documentId = await page.evaluate(() => {
      const scope = window.docsTest.editor.runtime.capture()
      scope.close()
      return scope.documentId
    })
    await app.evaluate(({ Menu }) => {
      const item = Menu.getApplicationMenu()
        ?.items.find((menu) => menu.label === 'File')
        ?.submenu?.items.find((entry) => entry.label === 'New Document')
      if (!item) throw new Error('The native New Document command is unavailable')
      item.click()
    })
    await expect(body).toHaveText('')
    await page.keyboard.press('Control+o')
    await expect
      .poll(() =>
        page.evaluate(() => {
          const scope = window.docsTest.editor.runtime.capture()
          scope.close()
          return scope.documentId
        }),
      )
      .not.toBe(documentId)
    await expect(body).toHaveText('Agent Word text')
    const lateOutput = await mcp.callTool({
      name: 'replace_selection',
      arguments: { html: 'Late output' },
    })
    expect(lateOutput.isError).toBe(true)
    expect(lateOutput.structuredContent).toMatchObject({
      errorCode: 'stale_target',
      committed: false,
    })
    await expect(body).toHaveText('Agent Word text')
    await mcp.close()
    await documentBridge.close()
    const reopenedBinding = randomUUID()
    await page.evaluate((binding) => {
      window.docsTest.editor.commands.bindScope(binding, window.docsTest.editor.runtime.capture())
    }, reopenedBinding)
    documentBridge = await createBridge(reopenedBinding)
    mcp = new Client({ name: 'reopened-document-tools', version: '1' }, { capabilities: {} })
    await mcp.connect(
      new StdioClientTransport({
        command: process.execPath,
        args: [resolve(__dirname, '../apps/docs/out/services/mcp-stdio.cjs')],
        env: {
          ...Object.fromEntries(
            Object.entries(process.env).filter(
              (entry): entry is [string, string] => entry[1] !== undefined,
            ),
          ),
          FORAOFFICE_DOCS_ENDPOINT: documentBridge.url,
          FORAOFFICE_DOCS_TOKEN: documentBridge.token,
        },
      }),
    )
    await tool('read_document', {})
    await tool('get_document_context', {})
    expect(await toolOutput('read_blocks', { startBlockIndex: 0, endBlockIndex: 0 })).toContain(
      'Agent',
    )
    const tracking = page.locator('.status-track')
    await tracking.click()
    await expect(tracking).toHaveAttribute('aria-pressed', 'true')
    const stalePolicy = await mcp.callTool({
      name: 'replace_blocks',
      arguments: {
        startBlockIndex: 0,
        endBlockIndex: 0,
        html: '<p>Rejected runtime revision</p>',
      },
    })
    expect(stalePolicy.structuredContent).toMatchObject({
      errorCode: 'stale_target',
      committed: false,
    })
    await tool('get_document_context', {})
    await tool('replace_blocks', {
      startBlockIndex: 0,
      endBlockIndex: 0,
      html: '<p>Rejected runtime revision</p>',
    })
    expect(await toolOutput('read_revisions')).toContain('Rejected runtime revision')
    expect(await toolOutput('reject_changes', { all: true })).toMatch(
      /Rejected [1-9]\d* tracked change/,
    )
    await expect(body).toHaveText('Agent Word text')
    await tool('replace_blocks', {
      startBlockIndex: 0,
      endBlockIndex: 0,
      html: '<p>Agent Word text accepted</p>',
    })
    expect(await toolOutput('accept_changes', { all: true })).toMatch(
      /Accepted [1-9]\d* tracked change/,
    )
    await expect(body).toHaveText('Agent Word text accepted')
    await tracking.click()
    await expect(tracking).toHaveAttribute('aria-pressed', 'false')
    await tool('get_document_context', {})
    const comment = await toolOutput('add_comment', {
      blockIndex: 0,
      text: 'Agent',
      comment: 'Runtime comment',
    })
    const commentId = /Added comment (\S+) on block/.exec(comment)?.[1]
    expect(commentId, comment).toBeTruthy()
    await tool('reply_comment', { parentId: commentId, text: 'Runtime reply' })
    await tool('resolve_comment', { id: commentId })
    const resolvedComments = await toolOutput('read_comments')
    expect(resolvedComments).toContain('Runtime reply')
    expect(resolvedComments).toContain('[resolved]')
    const temporaryComment = await toolOutput('add_comment', {
      blockIndex: 0,
      text: 'Word',
      comment: 'Remove this comment',
    })
    const temporaryCommentId = /Added comment (\S+) on block/.exec(temporaryComment)?.[1]
    expect(temporaryCommentId, temporaryComment).toBeTruthy()
    await tool('delete_comment', { id: temporaryCommentId })
    expect(await toolOutput('read_comments')).not.toContain('Remove this comment')
    const footnote = await toolOutput('insert_footnote', {
      blockIndex: 0,
      afterText: 'Agent',
      text: 'Original footnote',
    })
    const footnoteId = /Inserted footnote (\S+) in block/.exec(footnote)?.[1]
    expect(footnoteId, footnote).toBeTruthy()
    await tool('edit_note', {
      kind: 'footnote',
      id: footnoteId,
      find: 'Original',
      replace: 'Runtime',
    })
    const temporaryNote = await toolOutput('insert_endnote', {
      blockIndex: 0,
      afterText: 'Word',
      text: 'Remove this note',
    })
    const temporaryNoteId = /Inserted endnote (\S+) in block/.exec(temporaryNote)?.[1]
    expect(temporaryNoteId, temporaryNote).toBeTruthy()
    await tool('delete_note', { kind: 'endnote', id: temporaryNoteId })
    expect(await toolOutput('read_notes')).not.toContain('Remove this note')
    await tool('insert_endnote', { blockIndex: 0, afterText: 'Word', text: 'Runtime endnote' })
    expect(await toolOutput('read_notes')).toContain('Runtime footnote')
    const calls = [
      { name: 'set_header_footer', input: { kind: 'header', text: 'Runtime header' } },
      { name: 'set_header_footer', input: { kind: 'footer', text: 'Runtime footer' } },
      { name: 'set_page_setup', input: { orientation: 'landscape' } },
      {
        name: 'define_style',
        input: { styleId: 'RuntimeStyle', name: 'Runtime Style', run: { bold: true } },
      },
      {
        name: 'apply_ops',
        input: {
          ops: [{ op: 'applyStyle', target: { blockIndexes: [0] }, styleId: 'RuntimeStyle' }],
        },
      },
      { name: 'set_watermark', input: { text: 'Runtime watermark' } },
      {
        name: 'insert_picture',
        input: {
          url: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6f1cAAAAASUVORK5CYII=',
          width: '40px',
          height: '30px',
          float: { x: '10px', y: '12px', wrap: 'square' },
        },
      },
    ]
    for (const call of calls) await tool(call.name, call.input)
    expect(await toolOutput('list_styles')).toContain('RuntimeStyle')
    await tool('insert_content', { html: '<p>Runtime second block</p>', afterBlockIndex: 0 })
    await tool('insert_chart', {
      kind: 'bar',
      title: 'Original chart',
      categories: ['A', 'B'],
      series: [{ name: 'Values', values: [1, 2] }],
      afterBlockIndex: 1,
    })
    await tool('edit_chart', {
      blockIndex: 2,
      title: 'Runtime chart',
      series: [{ index: 0, values: [3, 4] }],
    })
    await tool('insert_section_break', { afterBlockIndex: 1, type: 'continuous' })
    await tool('insert_text_box', {
      text: 'Runtime text box',
      width: '96px',
      height: '48px',
      x: '5px',
      y: '5px',
      afterBlockIndex: 0,
    })
    const map = await tool('generate_mindmap', {
      root: { label: 'Runtime map', children: [{ label: 'Word text', block: 0 }] },
    })
    expect(map._meta?.['io.foraoffice/display']).toMatchObject({ kind: 'mindmap' })
    // Cursor-relative image insertion targets the original captured paragraph at index zero.
    await tool('insert_image', {
      url: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6f1cAAAAASUVORK5CYII=',
    })
    const media = await tool('read_document_media', { blockIndex: 0 })
    expect(media.content).toContainEqual({
      type: 'image',
      data: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6f1cAAAAASUVORK5CYII=',
      mimeType: 'image/png',
    })
    await tool('save_document', { path: output, overwrite: true })
    const stores = await JSZip.loadAsync(await readFile(output))
    expect(await stores.file('word/comments.xml')!.async('string')).toContain('Runtime comment')
    expect(await stores.file('word/comments.xml')!.async('string')).toContain('Runtime reply')
    expect(await stores.file('word/comments.xml')!.async('string')).not.toContain(
      'Remove this comment',
    )
    expect(await stores.file('word/footnotes.xml')!.async('string')).toContain('Runtime footnote')
    expect(await stores.file('word/endnotes.xml')!.async('string')).toContain('Runtime endnote')
    expect(await stores.file('word/styles.xml')!.async('string')).toContain('RuntimeStyle')
    const savedBody = await stores.file('word/document.xml')!.async('string')
    expect(savedBody).toContain('RuntimeStyle')
    expect(savedBody).toContain('landscape')
    expect(savedBody).toContain('Runtime text box')
    expect(savedBody).toContain('continuous')
    expect(savedBody).not.toContain('Rejected runtime revision')
    const chartXml = await Promise.all(
      Object.values(stores.files)
        .filter((file) => /^word\/charts\/chart\d+\.xml$/.test(file.name))
        .map((file) => file.async('string')),
    )
    expect(chartXml.join('\n')).toContain('Runtime chart')
    const headerFooter = await Promise.all(
      Object.values(stores.files)
        .filter((file) => /^word\/(header|footer).*\.xml$/.test(file.name))
        .map((file) => file.async('string')),
    )
    expect(headerFooter.join('\n')).toContain('Runtime header')
    expect(headerFooter.join('\n')).toContain('Runtime footer')
    expect(headerFooter.join('\n')).toContain('Runtime watermark')
    const images = Object.values(stores.files).filter((file) =>
      /^word\/media\/.*\.png$/.test(file.name),
    )
    expect(images.length).toBeGreaterThanOrEqual(1)
    for (const image of images)
      expect(await image.async('base64')).toBe(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6f1cAAAAASUVORK5CYII=',
      )
    expect(savedBody).toContain('wp:anchor')
    const originalDocumentText = await body.textContent()
    await app.evaluate(({ shell }) => {
      shell.showItemInFolder = () => {}
    })
    await tool('get_document_context', {})
    for (const [type, content] of [
      ['md', '# Created Markdown\n\nSupplied content.'],
      ['html', '<!doctype html><html><body>Supplied HTML.</body></html>'],
    ] as const) {
      await tool('create_document', { type, title: `Created ${type}`, content })
      expect(await readFile(join(documents, `Created ${type}.${type}`), 'utf8')).toBe(content)
    }
    await tool('create_document', {
      type: 'pdf',
      title: 'Created PDF',
      content: '<p>Supplied PDF content.</p>',
    })
    const createdPdfPath = join(documents, 'Created PDF.pdf')
    const createdPdf = await readFile(createdPdfPath)
    expect(createdPdf.subarray(0, 5).toString()).toBe('%PDF-')
    const pdfText = await page.evaluate(
      (path) => window.desktop.readAttachment(path, 0, 24_000),
      createdPdfPath,
    )
    expect(pdfText).toMatchObject({ ok: true })
    expect(pdfText.text).toContain('Supplied PDF content.')
    const newDocumentWindow = app.waitForEvent('window')
    await tool('create_document', {
      type: 'docx',
      title: 'Created Word',
      content: '<p>Supplied Word content.</p>',
    })
    const created = await newDocumentWindow
    await expect(created.locator('.ProseMirror').first()).toHaveText('Supplied Word content.')
    const createdWord = join(documents, 'Created Word.docx')
    await expect.poll(() => existsSync(createdWord)).toBe(true)
    expect(
      await (
        await JSZip.loadAsync(await readFile(createdWord))
      )
        .file('word/document.xml')!
        .async('string'),
    ).toContain('Supplied Word content.')
    await expect(body).toHaveText(originalDocumentText!)
    const deferred = new Set([
      'read_attachment',
      'search_knowledge_base',
      'read_knowledge_file',
      'run_cli',
    ])
    expect([...successfulTools].sort()).toEqual(
      [...LIVE_DOCS_TOOL_NAMES, 'save_document'].filter((name) => !deferred.has(name)).sort(),
    )
    await writeFile(
      join(scratch, 'mcp-tool-coverage.json'),
      JSON.stringify(
        { passed: [...successfulTools].sort(), notCoveredHere: [...deferred].sort() },
        null,
        2,
      ),
    )
    const protectedSave = await page.evaluate(async (path) => {
      await window.desktop.setDocPassword(path, 'isolated-test-password')
      const caller = window.docsTest.editor.runtime.connect({ id: 'protected-save-harness' })
      try {
        caller.context()
        return await caller.save({ path, overwrite: true })
      } finally {
        caller.close()
      }
    }, output)
    expect(protectedSave.ok).toBe(true)
    expect((await readFile(output)).subarray(0, 8).toString('hex')).toBe('d0cf11e0a1b11ae1')
    expect(
      await page.evaluate(
        async (path) => (await window.desktop.openDocxDecrypt(path, 'isolated-test-password')).ok,
        output,
      ),
    ).toBe(true)
    await page.screenshot({ path: join(scratch, 'editor.png') })
  } finally {
    await mcp?.close()
    await documentBridge?.close()
    files.dispose()
    if (app.process().exitCode === null)
      await app.evaluate(({ BrowserWindow }) => {
        for (const window of BrowserWindow.getAllWindows()) window.destroy()
      })
    await app.close()
    server.closeAllConnections()
    await new Promise<void>((done, reject) =>
      server.close((error) => (error ? reject(error) : done())),
    )
  }
})

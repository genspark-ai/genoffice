import { test, expect } from '@playwright/test'
import { launchSdkTestHost, sdkHostIsElectron, type SdkTestHost } from './sdk-host'
import { createServer } from 'node:http'
import { readFile, writeFile } from 'node:fs/promises'
import { extname, resolve, sep } from 'node:path'
import { buildDocx, TINY_PNG_BASE64 } from '../packages/docx-engine/tests/helpers/build-docx'
import { inlinePictureDrawingXml } from '../packages/docx-engine/src'
import JSZip from 'jszip'
import type { DocsEditorHandle } from '../apps/docs/src/renderer/embedded'
import type { DocsQueryResult } from '../apps/docs/src/shared/document-query'

type FixtureWindow = Window & { editor: DocsEditorHandle; sdkReady: boolean; saved: number[] }

test('embedded SDK preserves table-cell revisions and template structures through public editing', async () => {
  test.setTimeout(120_000)
  const bytes = await buildDocx({
    bodyXml:
      '<w:p><w:r><w:t>Condition report</w:t></w:r></w:p>' +
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t>Bridge has old cracks; inspect further.</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:p><w:r><w:t>Keep adjacent cell.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
      '<w:p><w:r><w:t>Keep closing paragraph.</w:t></w:r></w:p>',
  })
  const root = resolve(__dirname, '../apps/docs/out/embedded')
  const server = createServer(async (request, response) => {
    const path = new URL(request.url ?? '/', 'http://localhost').pathname
    if (path === '/') {
      response.setHeader('content-type', 'text/html')
      response.end(`<!doctype html><html><head><link rel="stylesheet" href="/sdk/editor.css"></head><body>
        <div id="root"></div><script type="module">
        import {mountDocsEditor,createExternalDocsHost} from '/sdk/editor.js';
        const defaults={getLanguage:'en',getSystemLocale:'en-US',getTheme:'light',getDocumentTheme:'follow',
          getAuthorName:'SDK reviewer',getAutoSaveDefault:{on:false,updatedAt:0},
          getAiPanelPrefs:{side:'left',fontSize:'default',spellcheck:false,openInNewDocs:false,showAiSidebar:false},
          spellLanguages:{active:[],available:[]},docPasswordIntentRevision:0};
        const services=new Proxy({}, {get(_,name){
          if(name==='then')return undefined;
          if(name.startsWith('on'))return ()=>()=>{};
          if(name==='confirmDocumentReplace')return async()=>true;
          return async()=>defaults[name];
        }});
        window.saved=[];
        const doc=()=>({id:'table:fixture',name:'Table.docx',data:window.saved.length
          ?Uint8Array.from(window.saved).buffer
          :Uint8Array.from(atob('${Buffer.from(bytes).toString('base64')}'),c=>c.charCodeAt(0)).buffer});
        const host=createExternalDocsHost('${sdkHostIsElectron ? 'electron' : 'browser'}',services);
        host.documents={initialDocument:async()=>doc(),openDocument:async()=>doc(),
          saveDocument:async r=>{window.saved=Array.from(new Uint8Array(r.data));return {ok:true,committed:true,pending:false};}};
        window.editor=await mountDocsEditor(document.getElementById('root'),{host});
        await window.editor.ready;window.sdkReady=true;
        </script></body></html>`)
      return
    }
    const file = resolve(root, decodeURIComponent(path.slice(5)))
    if (!path.startsWith('/sdk/') || !file.startsWith(root + sep)) {
      response.writeHead(404).end()
      return
    }
    try {
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
  if (!address || typeof address === 'string') throw new Error('No SDK fixture port')
  let browser: SdkTestHost | undefined
  try {
    browser = await launchSdkTestHost()
    const page = await browser.newPage({ viewport: { width: 1360, height: 1000 } })
    await page.goto(`http://127.0.0.1:${address.port}/`)
    await page.waitForFunction(() => (window as unknown as FixtureWindow).sdkReady)
    await page.getByRole('button', { name: 'Review', exact: true }).click()
    const track = page.getByRole('button', { name: 'Track Changes', exact: true })
    await track.click()
    await expect(track).toHaveClass(/active/)
    const replace = await page.evaluate(async () => {
      const { editor } = window as unknown as FixtureWindow
      const client = editor.runtime.connect({ id: 'table-replacement' })
      try {
        client.context()
        return await client.execute({
          id: 'replace',
          name: 'replace_blocks',
          input: {
            startBlockIndex: 1,
            endBlockIndex: 1,
            html: '<table><tr><td>Bridge has fine cracks; inspect further.</td><td>Keep adjacent cell.</td></tr></table>',
          },
        })
      } finally {
        client.close()
      }
    })
    expect(replace.isError, replace.output).not.toBe(true)
    const body = page.locator('.ProseMirror[contenteditable="true"]').first()
    const oldTable = body.locator('table.doc-del')
    const newTable = body.locator('table.doc-ins')
    for (const mode of ['No Markup', 'Original', 'Simple Markup', 'All Markup']) {
      await page.getByRole('button', { name: 'Markup', exact: true }).click()
      await page
        .locator('.layout-menu')
        .getByRole('button', { name: new RegExp(`^${mode}`) })
        .click()
      if (mode === 'No Markup' || mode === 'Simple Markup') {
        await expect(oldTable).toBeHidden()
        await expect(newTable).toBeVisible()
      } else if (mode === 'Original') {
        await expect(oldTable).toBeVisible()
        await expect(newTable).toBeHidden()
      } else {
        await expect(oldTable).toBeVisible()
        await expect(newTable).toBeVisible()
      }
    }
    // Use the browser's selection machinery so ProseMirror must observe a
    // genuine DOM selection inside the rendered, still-pending new table.
    await newTable
      .locator('td')
      .first()
      .evaluate((cell) => {
        const walker = document.createTreeWalker(cell, NodeFilter.SHOW_TEXT)
        let node: Node | null
        while ((node = walker.nextNode())) {
          const at = node.textContent!.indexOf('fine cracks')
          if (at < 0) continue
          const range = document.createRange()
          range.setStart(node, at)
          range.setEnd(node, at + 'fine cracks'.length)
          const selection = window.getSelection()!
          selection.removeAllRanges()
          selection.addRange(range)
          document.dispatchEvent(new Event('selectionchange'))
          return
        }
        throw new Error('Cell text was not rendered')
      })
    await expect
      .poll(() =>
        page.evaluate(() => {
          const { editor } = window as unknown as FixtureWindow
          const client = editor.runtime.connect({ id: crypto.randomUUID() })
          try {
            return client.context().text
          } finally {
            client.close()
          }
        }),
      )
      .toContain('Selected text (11 characters): "fine cracks"')
    const context = await page.evaluate(() => {
      const { editor } = window as unknown as FixtureWindow
      const scope = editor.runtime.capture()
      editor.commands.bindScope('cell-polish', scope)
      return scope.text
    })
    expect(context).toContain('<sel>fine cracks</sel>')
    await page.getByRole('button', { name: 'Home', exact: true }).click()
    const result = await page.evaluate(async () => {
      const { editor } = window as unknown as FixtureWindow
      return editor.commands.handle({
        requestId: 'cell-polish',
        callerId: 'external-ai',
        bindingId: 'cell-polish',
        command: 'execute_tool',
        payload: { name: 'replace_selection', input: { html: 'minor surface cracks' } },
      }) as Promise<{ isError?: boolean; mutated?: boolean }>
    })
    expect(result.isError).not.toBe(true)
    expect(result.mutated).toBe(true)
    await expect(oldTable).toContainText('Bridge has old cracks; inspect further.')
    await expect(newTable.locator('td').nth(1)).toHaveText('Keep adjacent cell.')
    await expect(newTable.locator('span.doc-del')).toHaveText('fine cracks')
    await expect(newTable.locator('span.doc-ins')).toHaveText('minor surface cracks')
    const reloaded = await page.evaluate(async () => {
      const { editor } = window as unknown as FixtureWindow
      const saved = await editor.documents.save()
      const reopened = await editor.documents.open()
      const client = editor.runtime.connect({ id: 'after-reload' })
      try {
        client.context()
        const result = await client.execute({
          id: 'read',
          name: 'read_blocks',
          input: { startBlockIndex: 0, endBlockIndex: 3 },
        })
        return { saved, reopened, result }
      } finally {
        client.close()
      }
    })
    expect(reloaded.saved).toBe(true)
    expect(reloaded.reopened).toBe(true)
    expect(reloaded.result.isError).not.toBe(true)
    expect(reloaded.result.output).toContain('Bridge has minor surface cracks; inspect further.')
    expect(reloaded.result.output).toContain('Keep adjacent cell.')
    expect(reloaded.result.output).toContain('Keep closing paragraph.')
    expect(reloaded.result.output).not.toContain('old cracks')
    expect(reloaded.result.output).not.toContain('fine cracks')
    await page.screenshot({ path: test.info().outputPath('sdk-cell-polish.png') })

    const template = await buildDocx({
      withImage: true,
      bodyXml:
        '<w:p><w:r><w:t>Historical description</w:t></w:r></w:p>' +
        '<w:p/><w:p/>' +
        '<w:p><w:r><w:t>Before protected structures</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>' +
        '<w:p><w:fldSimple w:instr="DATE"><w:r><w:t>2020-01-01</w:t></w:r></w:fldSimple></w:p>' +
        `<w:p><w:r>${inlinePictureDrawingXml({ rId: 'rId10', docPrId: 1, widthPx: 100, heightPx: 80 })}</w:r></w:p>` +
        '<w:p><w:r><w:t>Final paragraph</w:t></w:r></w:p>',
    })
    expect(
      await page.evaluate(async (bytes) => {
        const fixture = window as unknown as FixtureWindow
        fixture.saved = bytes
        return fixture.editor.documents.open()
      }, Array.from(template)),
    ).toBe(true)
    // Use the same public storage and runtime as an embedding application.
    const contract = await page.evaluate(async () => {
      const { editor } = window as unknown as FixtureWindow
      const client = editor.runtime.connect({ id: 'template-contract' })
      const execute = (name: string, input: Record<string, unknown>) =>
        client.execute({ id: crypto.randomUUID(), name, input })
      try {
        const context = client.context().text
        const emptyBefore = await execute('read_blocks', { startBlockIndex: 0, endBlockIndex: 3 })
        const roundtrip = await execute('replace_blocks', {
          startBlockIndex: 0,
          endBlockIndex: 3,
          html: emptyBefore.output,
        })
        const emptyAfter = await execute('read_blocks', { startBlockIndex: 0, endBlockIndex: 3 })
        const before = await execute('read_blocks', { startBlockIndex: 0, endBlockIndex: 7 })
        const refused = await execute('replace_blocks', {
          startBlockIndex: 0,
          endBlockIndex: 7,
          html: before.output.replace('Historical description', '{{description}}'),
        })
        const after = await execute('read_blocks', { startBlockIndex: 0, endBlockIndex: 7 })
        return {
          context,
          emptyBefore,
          roundtrip,
          emptyAfter,
          before,
          refused,
          after,
          modules: performance
            .getEntriesByType('resource')
            .map((entry) => entry.name)
            .filter((name) => /App-|embedded-/.test(name)),
        }
      } finally {
        client.close()
      }
    })
    await writeFile(
      test.info().outputPath('sdk-template-contract.json'),
      JSON.stringify(contract, null, 2),
    )
    expect(contract.emptyBefore.isError, contract.emptyBefore.output).not.toBe(true)
    expect(contract.emptyBefore.output.match(/<p><\/p>/g)).toHaveLength(2)
    expect(contract.roundtrip.isError, contract.roundtrip.output).not.toBe(true)
    expect(
      contract.emptyAfter.output.match(/<p><\/p>/g) ?? [],
      contract.emptyAfter.output,
    ).toHaveLength(2)
    expect(contract.before.output).toContain('data-doc-readonly="field"')
    expect(contract.before.output.match(/data-doc-readonly=/g)).toHaveLength(3)
    expect(contract.refused).toMatchObject({ isError: true, mutated: false })
    expect(contract.refused.output).toMatch(/read-only/i)
    expect(contract.after.output).toBe(contract.before.output)

    const saveTemplate = () =>
      page.evaluate(async () => {
        const fixture = window as unknown as FixtureWindow
        if (!(await fixture.editor.documents.save())) throw new Error('Template save failed')
        return fixture.saved
      })
    const beforeEdit = await saveTemplate()
    const inPlace = await page.evaluate(async () => {
      const { editor } = window as unknown as FixtureWindow
      const client = editor.runtime.connect({ id: 'template-in-place' })
      try {
        client.context()
        return await client.execute({
          id: 'template-placeholder',
          name: 'apply_ops',
          input: {
            ops: [
              {
                op: 'findReplace',
                find: 'Historical description',
                replace: '{{description}}',
                target: { blockIndexes: [0] },
              },
            ],
          },
        })
      } finally {
        client.close()
      }
    })
    expect(inPlace.isError, inPlace.output).not.toBe(true)
    const afterEdit = await saveTemplate()
    const structure = async (bytes: number[]) => {
      const zip = await JSZip.loadAsync(Uint8Array.from(bytes))
      const xml = await zip.file('word/document.xml')!.async('string')
      return {
        sections: xml.match(/<w:sectPr[\s>][\s\S]*?<\/w:sectPr>/g),
        fields: xml.match(/<w:fldSimple[\s>][\s\S]*?<\/w:fldSimple>/g),
        drawings: xml.match(/<w:drawing[\s>][\s\S]*?<\/w:drawing>/g),
        image: await zip.file('word/media/image1.png')!.async('base64'),
      }
    }
    const originalStructure = await structure(beforeEdit)
    expect(originalStructure.sections).toHaveLength(2)
    expect(originalStructure.fields).toHaveLength(1)
    expect(originalStructure.drawings).toHaveLength(1)
    expect(await structure(afterEdit)).toEqual(originalStructure)
    const templateReload = await page.evaluate(async () => {
      const { editor } = window as unknown as FixtureWindow
      const reopened = await editor.documents.open()
      const client = editor.runtime.connect({ id: 'template-reopened' })
      try {
        client.context()
        return {
          reopened,
          read: await client.execute({
            id: 'read',
            name: 'read_blocks',
            input: { startBlockIndex: 0, endBlockIndex: 7 },
          }),
        }
      } finally {
        client.close()
      }
    })
    expect(templateReload.reopened).toBe(true)
    expect(templateReload.read.isError, templateReload.read.output).not.toBe(true)
    expect(templateReload.read.output).toContain('{{description}}')
    expect(templateReload.read.output.match(/<p><\/p>/g)).toHaveLength(2)
    expect(templateReload.read.output.match(/data-doc-readonly=/g)).toHaveLength(3)
    expect(await structure(await saveTemplate())).toEqual(originalStructure)

    const freshness = await page.evaluate(async () => {
      const { editor } = window as unknown as FixtureWindow
      const reader = editor.runtime.connect({ id: 'template-stale-reader' })
      const writer = editor.runtime.connect({ id: 'template-concurrent-writer' })
      const edit = {
        id: 'reader-edit',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>{{description}}</p>' },
      }
      const read = {
        id: 'reader-blocks',
        name: 'read_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0 },
      }
      try {
        reader.context()
        writer.context()
        const bodyEdit = await writer.execute({
          id: 'concurrent-body',
          name: 'apply_ops',
          input: {
            ops: [
              {
                op: 'findReplace',
                find: 'Final paragraph',
                replace: 'Final amended paragraph',
                target: { blockIndexes: [7] },
              },
            ],
          },
        })
        const bodyStale = await reader.execute(edit)
        const bodyRead = await reader.execute(read)
        const bodyRecovered = await reader.execute(edit)
        reader.context()
        writer.context()
        const comment = await writer.execute({
          id: 'concurrent-comment',
          name: 'add_comment',
          input: { blockIndex: 0, comment: 'Review the template placeholder.' },
        })
        const sideRead = await reader.execute(read)
        const sideStale = await reader.execute(edit)
        const fullRead = await reader.execute({
          id: 'reader-context',
          name: 'get_document_context',
          input: {},
        })
        const sideRecovered = await reader.execute(edit)
        return {
          bodyEdit,
          bodyStale,
          bodyRead,
          bodyRecovered,
          comment,
          sideRead,
          sideStale,
          fullRead,
          sideRecovered,
        }
      } finally {
        reader.close()
        writer.close()
      }
    })
    for (const result of [
      freshness.bodyEdit,
      freshness.bodyRead,
      freshness.bodyRecovered,
      freshness.comment,
      freshness.sideRead,
      freshness.fullRead,
      freshness.sideRecovered,
    ])
      expect(result.isError, result.output).not.toBe(true)
    expect(freshness.bodyStale).toMatchObject({
      isError: true,
      errorCode: 'stale_target',
      committed: false,
    })
    expect(freshness.bodyStale.output).toBe(
      'The document body changed since this caller read it; use read_blocks, read_document or get_document_context before retrying.',
    )
    expect(freshness.sideStale).toMatchObject({
      isError: true,
      errorCode: 'stale_target',
      committed: false,
    })
    expect(freshness.sideStale.output).toBe(
      'Document-side state changed since this caller read it; use read_document or get_document_context for fresh full context. read_blocks does not refresh document-side state.',
    )

    const imageTable = Array.from(
      await buildDocx({
        withImage: true,
        bodyXml:
          '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="5000"/></w:tblGrid>' +
          '<w:tr><w:tc><w:p><w:r><w:t>Old label</w:t></w:r></w:p></w:tc>' +
          `<w:tc><w:p><w:r>${inlinePictureDrawingXml({ rId: 'rId10', docPrId: 1, widthPx: 100, heightPx: 80 })}</w:r></w:p></w:tc></w:tr></w:tbl>`,
      }),
    )
    for (const keep of [false, true]) {
      const result = await page.evaluate(
        async ({ bytes, keep }) => {
          const fixture = window as unknown as FixtureWindow
          fixture.saved = bytes
          const { editor } = fixture
          const opened = await editor.documents.open()
          const client = editor.runtime.connect({ id: 'image-replacement' })
          try {
            client.context()
            const read = {
              id: 'read',
              name: 'read_blocks',
              input: { startBlockIndex: 0, endBlockIndex: 0 },
            }
            const before = await client.execute(read)
            const queried = await client.execute({
              id: 'query-table',
              name: 'query_document',
              input: { domain: 'body', type: 'docTable', startBlockIndex: 0, endBlockIndex: 0 },
            })
            if (queried.isError) throw new Error(queried.output)
            const query = JSON.parse(queried.output) as DocsQueryResult
            const table = query.items[0]
            if (!query.page.complete || !table?.complete || !table.structuresComplete)
              throw new Error('The table query must provide complete content and structural intent')
            const image = new DOMParser()
              .parseFromString(table.html, 'text/html')
              .querySelector('img')
            if (!image) throw new Error('Expected the public table image reference')
            const imageHtml = image.outerHTML
            const imageRef = image.getAttribute('src')
            const deleteStructures = table
              .structures!.filter(
                (structure) => structure.kind === 'table' || (!keep && structure.kind === 'image'),
              )
              .map((structure) => structure.id)
            const replacement = await client.execute({
              id: 'replace',
              name: 'replace_blocks',
              input: {
                startBlockIndex: 0,
                endBlockIndex: 0,
                html: `<table>${[0, 1, 2].map((i) => `<tr><td>Parameter ${i}</td><td>${keep && i === 0 ? imageHtml : ''}</td></tr>`).join('')}</table>`,
                deleteStructures,
              },
            })
            const after = await client.execute(read)
            const saved = await editor.documents.save()
            return {
              opened,
              before,
              query,
              table,
              imageRef,
              deleteStructures,
              replacement,
              after,
              saved,
              bytes: fixture.saved,
            }
          } finally {
            client.close()
          }
        },
        { bytes: imageTable, keep },
      )
      expect(result.opened).toBe(true)
      expect(result.before.output.match(/<img /g)).toHaveLength(1)
      expect(result.imageRef).toMatch(/^media:[^:]+:\d+:\d+\.\d+$/)
      expect(result.before.output).toContain(result.imageRef!)
      expect(result.query.page.complete).toBe(true)
      expect(result.table.structuresComplete).toBe(true)
      expect(
        result.table.structures!.filter((structure) => structure.kind === 'table'),
      ).toHaveLength(1)
      expect(
        result.table.structures!.filter((structure) => structure.kind === 'image'),
      ).toHaveLength(1)
      expect(result.deleteStructures).toHaveLength(keep ? 1 : 2)
      expect(result.deleteStructures).toContain(
        result.table.structures!.find((structure) => structure.kind === 'table')!.id,
      )
      const imageStructure = result.table.structures!.find(
        (structure) => structure.kind === 'image',
      )!.id
      if (keep) expect(result.deleteStructures).not.toContain(imageStructure)
      else expect(result.deleteStructures).toContain(imageStructure)
      expect(result.replacement.isError, result.replacement.output).not.toBe(true)
      expect(result.replacement).toMatchObject({
        committed: true,
        mutated: true,
        receipt: {
          outcome: 'applied',
          revisionBefore: result.query.revision,
          persistence: 'pending',
        },
      })
      expect(result.after.output.match(/<img /g) ?? []).toHaveLength(keep ? 1 : 0)
      expect(result.saved).toBe(true)
      const zip = await JSZip.loadAsync(Uint8Array.from(result.bytes))
      const xml = await zip.file('word/document.xml')!.async('string')
      expect(xml.match(/<w:drawing[\s>]/g) ?? []).toHaveLength(keep ? 1 : 0)
      await page.evaluate(async () => {
        await (window as unknown as FixtureWindow).editor.documents.open()
      })
      const reloaded = await page.evaluate(async () => {
        const client = (window as unknown as FixtureWindow).editor.runtime.connect({
          id: 'image-reopened',
        })
        try {
          client.context()
          const read = await client.execute({
            id: 'read',
            name: 'read_blocks',
            input: { startBlockIndex: 0, endBlockIndex: 0 },
          })
          const images = await client.execute({
            id: 'active-images',
            name: 'query_document',
            input: { domain: 'images' },
          })
          if (images.isError) throw new Error(images.output)
          const query = JSON.parse(images.output) as DocsQueryResult
          const imageRead = query.items[0]
            ? await client.execute({
                id: 'image-bytes',
                name: 'read_document_media',
                input: { mediaRef: query.items[0].mediaRef },
              })
            : null
          return { ...read, images: query, imageRead }
        } finally {
          client.close()
        }
      })
      expect(reloaded.isError, reloaded.output).not.toBe(true)
      expect(reloaded.output.match(/<img /g) ?? []).toHaveLength(keep ? 1 : 0)
      expect(reloaded.images.items).toHaveLength(keep ? 1 : 0)
      expect(reloaded.images.page.complete).toBe(true)
      if (keep)
        expect(reloaded.imageRead).toMatchObject({
          committed: false,
          media: { data: TINY_PNG_BASE64, mimeType: 'image/png' },
        })
    }
  } finally {
    await browser?.close()
    await new Promise<void>((done) => server.close(() => done()))
  }
})

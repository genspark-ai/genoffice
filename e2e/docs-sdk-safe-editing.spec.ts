// Modified for ForaOffice: verify safe edits through real mounted SDK hosts and hosted bytes.
import { test, expect, type Page } from '@playwright/test'
import { createServer } from 'node:http'
import { createHash } from 'node:crypto'
import { readFile, writeFile } from 'node:fs/promises'
import { extname, posix, resolve, sep } from 'node:path'
import JSZip from 'jszip'
import { XMLParser } from 'fast-xml-parser'
import type { DocsEditorHandle } from '../apps/docs/src/renderer/embedded'
import type { DocsQueryResult } from '../apps/docs/src/shared/document-query'
import { launchSdkTestHost, sdkHostIsElectron, type SdkTestHost } from './sdk-host'

type FixtureWindow = Window & {
  editor: DocsEditorHandle
  ready: boolean
  fixture: { saved: number[]; saves: number }
}

async function withHostedDocument(original: Buffer, run: (page: Page) => Promise<void>) {
  const root = resolve(__dirname, '../apps/docs/out/embedded')
  const server = createServer(async (request, response) => {
    const path = new URL(request.url ?? '/', 'http://localhost').pathname
    if (path === '/') {
      response.setHeader('content-type', 'text/html')
      response.end(`<!doctype html><html><head><link rel="stylesheet" href="/sdk/editor.css"></head><body>
        <div id="root"></div><script type="module">
        import {mountDocsEditor,createExternalDocsHost} from '/sdk/editor.js';
        const defaults={getLanguage:'en',getSystemLocale:'en-US',getTheme:'light',getDocumentTheme:'follow',
          getAuthorName:'Safe editing fixture',getAutoSaveDefault:{on:false,updatedAt:0},
          getAiPanelPrefs:{side:'left',fontSize:'default',spellcheck:false,openInNewDocs:false,showAiSidebar:false},
          spellLanguages:{active:[],available:[]},docPasswordIntentRevision:0};
        const services=new Proxy({}, {get(_,name){
          if(name==='then')return undefined;
          if(name.startsWith('on'))return ()=>()=>{};
          if(name==='confirmDocumentReplace')return async()=>true;
          return async()=>defaults[name];
        }});
        window.fixture={saved:[],saves:0};
        const host=createExternalDocsHost('${sdkHostIsElectron ? 'electron' : 'browser'}',services);
        const initial=Uint8Array.from(atob('${original.toString('base64')}'),c=>c.charCodeAt(0));
        const doc=()=>({id:'safe-editing:fixture',name:'Safe editing.docx',
          data:Uint8Array.from(window.fixture.saved.length?window.fixture.saved:initial).buffer});
        host.documents={initialDocument:async()=>doc(),openDocument:async()=>doc(),
          saveDocument:async request=>{window.fixture.saved=Array.from(new Uint8Array(request.data));window.fixture.saves++;
            return {ok:true,committed:true,pending:false,document:{id:'safe-editing:fixture',name:'Safe editing.docx'}};}};
        window.editor=await mountDocsEditor(document.getElementById('root'),{host});
        await window.editor.ready;window.ready=true;
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
  if (!address || typeof address === 'string') throw new Error('No SDK fixture port')
  let browser: SdkTestHost | undefined
  try {
    browser = await launchSdkTestHost()
    const page = await browser.newPage()
    const errors: string[] = []
    page.on('pageerror', (error) => errors.push(error.message))
    await page.goto(`http://127.0.0.1:${address.port}/`)
    await page.waitForFunction(() => (window as unknown as FixtureWindow).ready)
    await run(page)
    expect(errors).toEqual([])
  } finally {
    await browser?.close()
    await new Promise<void>((done) => server.close(() => done()))
  }
}

test('mounted SDK validates queried edits and saves body and header effects in their original containers', async () => {
  test.setTimeout(120_000)
  const original = await readFile(
    resolve(__dirname, '../apps/docs/tests/fixtures/foraoffice-header-textbox.docx'),
  )
  await withHostedDocument(original, async (page) => {
    const result = await page.evaluate(async () => {
      const { editor, fixture } = window as unknown as FixtureWindow
      const caller = editor.runtime.connect({ id: 'safe-editing-acceptance' })
      try {
        const before = caller.context()
        const dirtyBefore = editor.documents.state()!.dirty
        const invalidImage = await caller.execute({
          id: 'invalid-image',
          name: 'replace_blocks',
          input: {
            startBlockIndex: 0,
            endBlockIndex: 0,
            html: `<p>INVALID_IMAGE_MUST_NOT_APPEAR <img src="media:${before.documentId}:${before.revision}:99999.0"></p>`,
          },
        })
        const afterRefusal = caller.context()
        const dirtyAfterRefusal = editor.documents.state()!.dirty
        const query = await caller.execute({
          id: 'query-body',
          name: 'query_document',
          input: { containsText: 'HEADER FIXTURE', limit: 1 },
        })
        if (query.isError) throw new Error(query.output)
        const queried: DocsQueryResult = JSON.parse(query.output)
        if (queried.items.length !== 1) throw new Error('Expected one body target')
        const target = queried.items[0].target
        const applied = await caller.execute({
          id: 'edit-queried-body',
          name: 'apply_ops',
          input: {
            ops: [
              {
                op: 'findReplace',
                target: { engine: target },
                find: 'HEADER FIXTURE',
                replace: 'QUERY EDIT',
              },
            ],
          },
        })
        const refreshed = await caller.execute({
          id: 'query-edited-body',
          name: 'query_document',
          input: { containsText: 'QUERY EDIT', limit: 1 },
        })
        if (refreshed.isError) throw new Error(refreshed.output)
        const current: DocsQueryResult = JSON.parse(refreshed.output)
        if (current.items.length !== 1) throw new Error('Expected the edited body target')
        const noOp = await caller.execute({
          id: 'unchanged-queried-body',
          name: 'apply_ops',
          input: {
            ops: [
              {
                op: 'findReplace',
                target: { engine: current.items[0].target },
                find: 'QUERY EDIT',
                replace: 'QUERY EDIT',
              },
            ],
          },
        })
        const beforeStale = caller.context()
        const stale = await caller.execute({
          id: 'expired-body-target',
          name: 'apply_ops',
          input: {
            ops: [
              {
                op: 'findReplace',
                target: { engine: target },
                find: 'HEADER FIXTURE',
                replace: 'STALE_EDIT_MUST_NOT_APPEAR',
              },
            ],
          },
        })
        const afterStale = caller.context()
        const header = await caller.execute({
          id: 'header',
          name: 'set_header_footer',
          input: {
            kind: 'header',
            section: 0,
            textboxIndex: 0,
            text: 'SAFE_HEADER {PAGE} / {NUMPAGES}',
          },
        })
        const saved = await editor.documents.save()
        return {
          before,
          invalidImage,
          afterRefusal,
          dirtyBefore,
          dirtyAfterRefusal,
          queried,
          applied,
          current,
          noOp,
          beforeStale,
          stale,
          afterStale,
          header,
          saved,
          state: editor.documents.state(),
          bytes: fixture.saved,
          saves: fixture.saves,
        }
      } finally {
        caller.close()
      }
    })
    expect(result.before.text).toContain('OLD_HEADER')
    expect(result.invalidImage).toMatchObject({
      isError: true,
      committed: false,
      mutated: false,
      errorCode: 'invalid_input',
      revision: result.before.revision,
    })
    expect(result.invalidImage.imageReferenceErrors).toHaveLength(1)
    expect(result.afterRefusal).toEqual(result.before)
    expect(result.dirtyBefore).toBe(false)
    expect(result.dirtyAfterRefusal).toBe(false)
    expect(result.queried.page.complete).toBe(true)
    expect(result.queried.items[0].supportedOperations).toContain('findReplace')
    expect(result.applied).toMatchObject({
      committed: true,
      mutated: true,
      matchCount: 1,
      changeCount: 1,
      receipt: {
        outcome: 'applied',
        revisionBefore: result.queried.revision,
        persistence: 'pending',
      },
    })
    expect(result.applied.isError).not.toBe(true)
    expect(result.applied.revision).toBeGreaterThan(result.queried.revision)
    expect(result.noOp).toMatchObject({
      committed: false,
      mutated: false,
      matchCount: 1,
      changeCount: 0,
      revision: result.current.revision,
      receipt: {
        outcome: 'no-op',
        revisionBefore: result.current.revision,
        revisionAfter: result.current.revision,
        persistence: 'unchanged',
      },
    })
    expect(result.noOp.isError).not.toBe(true)
    expect(result.stale).toMatchObject({
      isError: true,
      errorCode: 'stale_target',
      committed: false,
      mutated: false,
      revision: result.current.revision,
      receipt: { outcome: 'rejected' },
    })
    expect(result.afterStale).toEqual(result.beforeStale)
    expect(result.header).toMatchObject({ committed: true })
    expect(result.header.isError).not.toBe(true)
    expect(result.saved).toBe(true)
    expect(result.saves).toBe(1)
    expect(result.state).toMatchObject({ path: null, dirty: false })

    const saved = Uint8Array.from(result.bytes)
    await writeFile(test.info().outputPath('hosted-safe.docx'), saved)
    const zip = await JSZip.loadAsync(saved)
    const body = await zip.file('word/document.xml')!.async('string')
    expect(body).toContain('QUERY EDIT')
    expect(body).not.toContain('HEADER FIXTURE')
    expect(body).not.toContain('STALE_EDIT_MUST_NOT_APPEAR')
    expect(body).not.toContain('INVALID_IMAGE_MUST_NOT_APPEAR')
    const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '' })
    const section = parser.parse(body)['w:document']['w:body']['w:sectPr']
    const references = [section['w:headerReference']].flat()
    const rId = references.find((reference) => reference['w:type'] === 'default')['r:id']
    const rels = parser.parse(await zip.file('word/_rels/document.xml.rels')!.async('string'))
    const target = rels.Relationships.Relationship.find(
      (rel: { Id: string }) => rel.Id === rId,
    ).Target
    const part = posix.normalize(target.startsWith('/') ? target.slice(1) : `word/${target}`)
    const xml = await zip.file(part)!.async('string')
    expect(xml).not.toContain('OLD_HEADER')
    const boxes = [...xml.matchAll(/<w:txbxContent\b[^>]*>([\s\S]*?)<\/w:txbxContent>/g)]
    expect(boxes).toHaveLength(2)
    for (const [, contents] of boxes) {
      expect(contents).toContain('SAFE_HEADER')
      expect(contents).toMatch(/<w:instrText\b[^>]*>\s*PAGE\b/)
      expect(contents).toMatch(/<w:instrText\b[^>]*>\s*NUMPAGES\b/)
    }
    expect(xml.match(/<mc:Choice\b/g)).toHaveLength(1)
    expect(xml.match(/<mc:Fallback\b/g)).toHaveLength(1)
    expect(xml).toContain('<wp:extent cx="1828800" cy="1828800"/>')
    expect(xml).toContain('width:2in;height:2in')
    const paragraphs = [parser.parse(xml)['w:hdr']['w:p']].flat()
    expect(paragraphs).toHaveLength(1)

    const reopened = await page.evaluate(async () => {
      const { editor } = window as unknown as FixtureWindow
      const opened = await editor.documents.open()
      const caller = editor.runtime.connect({ id: 'safe-editing-reopened' })
      try {
        return { opened, context: caller.context(), body: editor.inspection.snapshot().text }
      } finally {
        caller.close()
      }
    })
    expect(reopened.opened).toBe(true)
    expect(reopened.context.documentId).not.toBe(result.before.documentId)
    expect(reopened.context.text).toContain('SAFE_HEADER {PAGE} / {NUMPAGES}')
    expect(reopened.context.text).not.toContain('OLD_HEADER')
    expect(reopened.body).toContain('QUERY EDIT')
    expect(reopened.body).not.toContain('HEADER FIXTURE')
    expect(reopened.body).not.toContain('STALE_EDIT_MUST_NOT_APPEAR')
    expect(reopened.body).not.toContain('INVALID_IMAGE_MUST_NOT_APPEAR')
  })
})

test('explicit section deletion updates live routing, undo and saved section ownership', async () => {
  test.setTimeout(120_000)
  const fixture = await JSZip.loadAsync(
    await readFile(
      resolve(__dirname, '../apps/docs/tests/fixtures/foraoffice-header-textbox.docx'),
    ),
  )
  const source = await fixture.file('word/document.xml')!.async('string')
  fixture.file(
    'word/document.xml',
    source.replace(
      '<w:body>',
      '<w:body><w:p><w:pPr><w:mirrorIndents/><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="10000" w:h="15000"/></w:sectPr></w:pPr><w:r><w:t>SECTION END</w:t></w:r></w:p>',
    ),
  )
  await withHostedDocument(await fixture.generateAsync({ type: 'nodebuffer' }), async (page) => {
    const result = await page.evaluate(async () => {
      const { editor, fixture } = window as unknown as FixtureWindow
      const caller = editor.runtime.connect({ id: 'section-deletion-acceptance' })
      const query = async (id: string, input: Record<string, string | number>) => {
        const result = await caller.execute({ id, name: 'query_document', input })
        if (result.isError) throw new Error(result.output)
        return JSON.parse(result.output) as DocsQueryResult
      }
      try {
        const before = caller.context()
        const sectionsBefore = await query('sections-before', { domain: 'sections' })
        const headersBefore = await query('headers-before', {
          domain: 'headers',
          type: 'textbox',
          containsText: 'OLD_HEADER',
        })
        const body = await query('section-target', {
          domain: 'body',
          containsText: 'SECTION END',
        })
        if (body.items.length !== 1) throw new Error('Expected one section-ending paragraph')
        const item = body.items[0]
        if (!item.structuresComplete) throw new Error('Expected the complete structure inventory')
        const boundaries = item.structures?.filter((structure) => structure.kind === 'section')
        if (boundaries?.length !== 1) throw new Error('Expected one section boundary')
        const input = {
          startBlockIndex: item.blockIndex,
          endBlockIndex: item.blockIndex,
          html: '<p>SECTION REMOVED</p>',
        }
        const refused = await caller.execute({
          id: 'ordinary-section-edit',
          name: 'replace_blocks',
          input,
        })
        const afterRefusal = caller.context()
        const removed = await caller.execute({
          id: 'explicit-section-edit',
          name: 'replace_blocks',
          input: { ...input, deleteStructures: [boundaries[0].id] },
        })
        const sectionsAfter = await query('sections-after', { domain: 'sections' })
        const headersAfter = await query('headers-after', {
          domain: 'headers',
          type: 'textbox',
          containsText: 'OLD_HEADER',
        })
        caller.context()
        const undo = await caller.execute({
          id: 'undo-section-delete',
          name: 'document_history',
          input: { action: 'undo' },
        })
        const sectionsUndone = await query('sections-undone', { domain: 'sections' })
        const headersUndone = await query('headers-undone', {
          domain: 'headers',
          type: 'textbox',
          containsText: 'OLD_HEADER',
        })
        caller.context()
        const redo = await caller.execute({
          id: 'redo-section-delete',
          name: 'document_history',
          input: { action: 'redo' },
        })
        const sectionsRedone = await query('sections-redone', { domain: 'sections' })
        caller.context()
        const routed = await caller.execute({
          id: 'edit-routed-header',
          name: 'set_header_footer',
          input: {
            kind: 'header',
            section: 0,
            textboxIndex: 0,
            text: 'ROUTED_HEADER {PAGE}/{NUMPAGES}',
          },
        })
        const headersRouted = await query('headers-routed', {
          domain: 'headers',
          type: 'textbox',
          containsText: 'ROUTED_HEADER',
        })
        const saved = await editor.documents.save()
        return {
          before,
          sectionsBefore,
          headersBefore,
          refused,
          afterRefusal,
          removed,
          sectionsAfter,
          headersAfter,
          undo,
          sectionsUndone,
          headersUndone,
          redo,
          sectionsRedone,
          routed,
          headersRouted,
          saved,
          bytes: fixture.saved,
        }
      } finally {
        caller.close()
      }
    })
    expect(result.sectionsBefore.items).toHaveLength(2)
    expect(result.headersBefore.items).toHaveLength(1)
    expect(result.headersBefore.items[0].target.address?.section).toBe(1)
    expect(result.refused).toMatchObject({ isError: true, committed: false, mutated: false })
    expect(result.afterRefusal).toEqual(result.before)
    expect(result.removed.isError).not.toBe(true)
    expect(result.removed).toMatchObject({ committed: true, mutated: true })
    expect(result.sectionsAfter.items).toHaveLength(1)
    expect(result.sectionsAfter.page.complete).toBe(true)
    expect(result.headersAfter.items).toHaveLength(1)
    expect(result.headersAfter.items[0].target.address?.section).toBe(0)
    expect(result.undo.isError).not.toBe(true)
    expect(result.sectionsUndone.items).toHaveLength(2)
    expect(result.headersUndone.items).toHaveLength(1)
    expect(result.headersUndone.items[0].target.address?.section).toBe(1)
    expect(result.redo.isError).not.toBe(true)
    expect(result.sectionsRedone.items).toHaveLength(1)
    expect(result.routed.isError).not.toBe(true)
    expect(result.routed).toMatchObject({ committed: true, mutated: true })
    expect(result.headersRouted.items).toHaveLength(1)
    expect(result.headersRouted.items[0].target.address?.section).toBe(0)
    expect(result.headersRouted.items[0].html).toContain('ROUTED_HEADER')
    expect(result.headersRouted.items[0].html).not.toContain('OLD_HEADER')
    expect(result.saved).toBe(true)

    const saved = Uint8Array.from(result.bytes)
    await writeFile(test.info().outputPath('hosted-section-delete.docx'), saved)
    const zip = await JSZip.loadAsync(saved)
    const xml = await zip.file('word/document.xml')!.async('string')
    expect(xml.match(/<w:sectPr\b/g)).toHaveLength(1)
    expect(xml).not.toContain('w:w="10000"')
    const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '' })
    const body = parser.parse(xml)['w:document']['w:body']
    const first = [body['w:p']].flat()[0]
    expect(first['w:pPr']).toHaveProperty('w:mirrorIndents')
    expect(first['w:pPr']).not.toHaveProperty('w:sectPr')
    const firstText = first['w:r']['w:t']
    expect(
      typeof firstText === 'object' && firstText !== null ? firstText['#text'] : firstText,
    ).toBe('SECTION REMOVED')
    const section = body['w:sectPr']
    expect(section['w:pgSz']['w:w']).toBe('11906')
    const reference = [section['w:headerReference']]
      .flat()
      .find((ref) => ref['w:type'] === 'default')
    const rels = parser.parse(await zip.file('word/_rels/document.xml.rels')!.async('string'))
    const relationship = [rels.Relationships.Relationship]
      .flat()
      .find((rel) => rel.Id === reference['r:id'])
    const target = relationship.Target
    const part = posix.normalize(target.startsWith('/') ? target.slice(1) : `word/${target}`)
    expect(zip.file(part)).not.toBeNull()
    const header = await zip.file(part)!.async('string')
    expect(header).toContain('ROUTED_HEADER')
    expect(header).not.toContain('OLD_HEADER')
    const originalSection = parser.parse(source)['w:document']['w:body']['w:sectPr']
    const originalReference = [originalSection['w:headerReference']]
      .flat()
      .find((ref) => ref['w:type'] === 'default')
    const originalRels = parser.parse(
      await fixture.file('word/_rels/document.xml.rels')!.async('string'),
    )
    const originalRelationship = [originalRels.Relationships.Relationship]
      .flat()
      .find((rel) => rel.Id === originalReference['r:id'])
    const originalTarget = originalRelationship.Target
    const originalPart = posix.normalize(
      originalTarget.startsWith('/') ? originalTarget.slice(1) : `word/${originalTarget}`,
    )
    const originalHeader = await fixture.file(originalPart)!.async('string')
    expect(relationship.Type).toBe(originalRelationship.Type)
    expect(part).not.toBe(originalPart)
    expect(await zip.file(originalPart)!.async('string')).toBe(originalHeader)
    const geometry = (xml: string) =>
      [...xml.matchAll(/<(wp:anchor|wp:inline|v:shape)\b[\s\S]*?<\/\1>/g)].map(([container]) =>
        createHash('sha256')
          .update(
            container.replace(
              /<w:txbxContent\b[^>]*>[\s\S]*?<\/w:txbxContent>/g,
              '<w:txbxContent/>',
            ),
          )
          .digest('hex'),
      )
    const originalGeometry = geometry(originalHeader)
    expect(originalGeometry).toHaveLength(2)
    expect(geometry(header)).toEqual(originalGeometry)
    const boxes = [...header.matchAll(/<w:txbxContent\b[^>]*>([\s\S]*?)<\/w:txbxContent>/g)]
    expect(boxes).toHaveLength(2)
    for (const [, contents] of boxes) {
      expect(contents).toContain('ROUTED_HEADER')
      expect(contents).toMatch(/<w:instrText\b[^>]*>\s*PAGE\b/)
      expect(contents).toMatch(/<w:instrText\b[^>]*>\s*NUMPAGES\b/)
      for (const kind of ['begin', 'separate', 'end'])
        expect(
          contents.match(new RegExp(`<w:fldChar\\b[^>]*w:fldCharType="${kind}"`, 'g')),
        ).toHaveLength(2)
    }
    expect(header.match(/<mc:Choice\b/g)).toHaveLength(1)
    expect(header.match(/<mc:Fallback\b/g)).toHaveLength(1)

    const reopened = await page.evaluate(async () => {
      const { editor } = window as unknown as FixtureWindow
      const opened = await editor.documents.open()
      const caller = editor.runtime.connect({ id: 'section-deletion-reopened' })
      try {
        const sections = await caller.execute({
          id: 'reopened-sections',
          name: 'query_document',
          input: { domain: 'sections' },
        })
        const headers = await caller.execute({
          id: 'reopened-headers',
          name: 'query_document',
          input: { domain: 'headers', type: 'textbox', containsText: 'ROUTED_HEADER' },
        })
        if (sections.isError || headers.isError) throw new Error(sections.output + headers.output)
        return {
          opened,
          sections: JSON.parse(sections.output) as DocsQueryResult,
          headers: JSON.parse(headers.output) as DocsQueryResult,
        }
      } finally {
        caller.close()
      }
    })
    expect(reopened.opened).toBe(true)
    expect(reopened.sections.items).toHaveLength(1)
    expect(reopened.headers.items).toHaveLength(1)
    expect(reopened.headers.items[0].target.address?.section).toBe(0)
  })
})

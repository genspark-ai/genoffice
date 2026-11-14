import { afterEach, expect, it } from 'vitest'
import { readFile } from 'node:fs/promises'
import { createHash } from 'node:crypto'
import { resolve } from 'node:path'
import { Editor } from '@tiptap/core'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import JSZip from 'jszip'
import {
  parseDocx,
  readSections,
  saveDocx,
  PAGE_MARK,
  TOTAL_PAGES_MARK,
} from '@genoffice/docx-engine'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { createHeaderFooterAccess } from '../src/renderer/ai/header-footer-access'
import { hfSaveOptions, type HfSectionState } from '../src/renderer/hf-sections'
import { createDocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { createDocsCommandRouter } from '../src/renderer/runtime/command-router'
import { createDocsMcpServer } from '../src/services/docs-mcp'

const editors: Editor[] = []
afterEach(() => editors.splice(0).forEach((editor) => editor.destroy()))

async function open(bytes: Uint8Array) {
  const parsed = await parseDocx(bytes)
  let state: HfSectionState = { sections: readSections(parsed), edits: {}, links: {}, parsed }
  let firstEnabled = 0
  let evenEnabled = 0
  const headerFooter = createHeaderFooterAccess({
    current: () => state,
    commit: (next) => {
      state = next
    },
    locked: () => false,
    evenOdd: () => false,
    enableFirst: () => {
      firstEnabled++
      return null
    },
    enableEven: () => {
      evenEnabled++
    },
  })
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  editors.push(editor)
  const runtime = createDocsRuntime({
    current: () => ({
      editor,
      identity: editor,
      stateVersion: state,
      numIds: { bullet: null, ordered: null },
      headerFooter,
    }),
  })
  const caller = runtime.connect({ id: 'textbox-header-write' })
  caller.context()
  const saved = async (section = 0) => {
    const bytes = await saveDocx(
      parsed,
      parsed.blocks
        .filter((block) => !block.hidden)
        .map((block) => ({
          kind: 'original' as const,
          docxIndex: block.docxIndex!,
        })),
      hfSaveOptions(state),
    )
    const reopened = await parseDocx(bytes)
    const zip = await JSZip.loadAsync(bytes)
    const rId = readSections(reopened)[section].headerRefs.default!
    const rels = new DOMParser().parseFromString(
      await zip.file('word/_rels/document.xml.rels')!.async('string'),
      'application/xml',
    )
    const target = [...rels.getElementsByTagName('Relationship')]
      .find((rel) => rel.getAttribute('Id') === rId)!
      .getAttribute('Target')!
    const path = target.startsWith('/') ? target.slice(1) : `word/${target}`
    return { part: reopened.hfParts![rId], xml: await zip.file(path)!.async('string'), reopened }
  }
  return {
    caller,
    runtime,
    saved,
    headerFooter,
    effects: () => ({ state, firstEnabled, evenEnabled }),
  }
}

it('replaces the sanitized header textbox in place through save and reopen', async () => {
  const bytes = await readFile(resolve(__dirname, 'fixtures/foraoffice-header-textbox.docx'))
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(
    '9a235ff2d09b3b17fba9f3a3ed39e86d6994da399e055256b894863362642625',
  )
  const { caller, runtime, saved } = await open(bytes)
  try {
    expect(caller.context().text).toContain('OLD_HEADER')
    expect(caller.context().text).toContain('"editable":true,"lifetime":"current-part"')
    const result = await caller.execute({
      id: 'replace',
      name: 'set_header_footer',
      input: {
        kind: 'header',
        section: 0,
        view: 'default',
        text: 'NEW_HEADER {PAGE} / {NUMPAGES}',
      },
    })
    expect(result).toMatchObject({ committed: true })
    expect(result.isError).not.toBe(true)
    const after = await saved()
    expect(after.part.text).toContain(`NEW_HEADER ${PAGE_MARK} / ${TOTAL_PAGES_MARK}`)
    expect(after.xml).not.toContain('OLD_HEADER')
    const xml = new DOMParser().parseFromString(after.xml, 'application/xml')
    const choices = [...xml.getElementsByTagName('mc:Choice')]
    const fallbacks = [...xml.getElementsByTagName('mc:Fallback')]
    expect(choices).toHaveLength(1)
    expect(fallbacks).toHaveLength(1)
    for (const branch of [...choices, ...fallbacks]) {
      const textboxes = [...branch.getElementsByTagName('w:txbxContent')]
      expect(textboxes).toHaveLength(1)
      expect(textboxes[0].textContent).toContain('NEW_HEADER')
      expect(textboxes[0].textContent).toContain('NUMPAGES')
    }
    expect(xml.getElementsByTagName('wp:anchor')).toHaveLength(1)
    expect(xml.getElementsByTagName('wp:extent')[0].getAttribute('cx')).toBe('1828800')
    // The containing paragraph remains only an anchor, with no new flow-text twin.
    const root = xml.documentElement
    expect([...root.children].filter((child) => child.tagName === 'w:p')).toHaveLength(1)
  } finally {
    runtime.dispose()
  }
})

it('updates a textbox compatibility fallback that carries namespace attributes', async () => {
  const zip = await JSZip.loadAsync(
    await readFile(resolve(__dirname, 'fixtures/foraoffice-header-textbox.docx')),
  )
  zip.file(
    'word/header1.xml',
    (await zip.file('word/header1.xml')!.async('string')).replace(
      '<mc:Fallback>',
      '<mc:Fallback xmlns:fixture="urn:fixture">',
    ),
  )
  const { caller, runtime, saved } = await open(await zip.generateAsync({ type: 'uint8array' }))
  try {
    expect(
      await caller.execute({
        id: 'namespace-fallback',
        name: 'set_header_footer',
        input: { kind: 'header', text: 'NEXT_HEADER {PAGE}' },
      }),
    ).toMatchObject({ committed: true })
    expect((await saved()).xml).not.toContain('OLD_HEADER')
  } finally {
    runtime.dispose()
  }
})

it('targets one textbox while keeping ordinary header text and other boxes unchanged', async () => {
  const bytes = await readFile(resolve(__dirname, 'fixtures/foraoffice-header-textbox.docx'))
  const zip = await JSZip.loadAsync(bytes)
  const original = await zip.file('word/header1.xml')!.async('string')
  const firstParagraph = original.slice(original.indexOf('<w:p '), original.lastIndexOf('</w:hdr>'))
  zip.file(
    'word/header1.xml',
    original.replace(
      firstParagraph,
      '<w:p><w:r><w:t>KEEP_FLOW</w:t></w:r></w:p>' +
        firstParagraph +
        firstParagraph.replaceAll('OLD_HEADER', 'OTHER_BOX'),
    ),
  )
  const { caller, runtime, saved } = await open(await zip.generateAsync({ type: 'uint8array' }))
  try {
    const result = await caller.execute({
      id: 'one-box',
      name: 'set_header_footer',
      input: {
        kind: 'header',
        textboxIndex: 1,
        text: 'CHANGED_BOX {PAGE}',
      },
    })
    expect(result.isError, result.output).not.toBe(true)
    const after = await saved()
    expect(after.part.text).toContain('KEEP_FLOW')
    expect(after.part.text).toContain('OLD_HEADER')
    expect(after.xml).not.toContain('OTHER_BOX')
    expect(after.part.text).toContain(`CHANGED_BOX ${PAGE_MARK}`)
    expect(
      after.part.paras
        .filter((para) => para.box)
        .map((para) => para.runs.map((r) => r.text).join('')),
    ).toHaveLength(2)
  } finally {
    runtime.dispose()
  }
})

it('clears textbox text and its old page fields without appending a flow paragraph', async () => {
  const { caller, runtime, saved } = await open(
    await readFile(resolve(__dirname, 'fixtures/foraoffice-header-textbox.docx')),
  )
  try {
    expect(
      await caller.execute({
        id: 'clear',
        name: 'set_header_footer',
        input: { kind: 'header', text: '' },
      }),
    ).toMatchObject({ committed: true })
    const after = await saved()
    expect(after.part.text).toBe('')
    expect(after.part.hasPageNumber).toBe(false)
    expect(after.xml).not.toContain('OLD_HEADER')
    const xml = new DOMParser().parseFromString(after.xml, 'application/xml')
    expect(
      [...xml.documentElement.children].filter((child) => child.tagName === 'w:p'),
    ).toHaveLength(1)
    expect(xml.getElementsByTagName('wp:anchor')).toHaveLength(1)
  } finally {
    runtime.dispose()
  }
})

it('keeps all replacement lines in the same textbox and isolates the edited section', async () => {
  const zip = await JSZip.loadAsync(
    await readFile(resolve(__dirname, 'fixtures/foraoffice-header-textbox.docx')),
  )
  const body = await zip.file('word/document.xml')!.async('string')
  const section = /<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/.exec(body)![0]
  zip.file(
    'word/document.xml',
    body.replace(
      '<w:body>',
      '<w:body>' + `<w:p><w:pPr>${section}</w:pPr><w:r><w:t>FIRST SECTION</w:t></w:r></w:p>`,
    ),
  )
  const { caller, runtime, saved } = await open(await zip.generateAsync({ type: 'uint8array' }))
  try {
    expect(
      await caller.execute({
        id: 'multiline',
        name: 'set_header_footer',
        input: {
          kind: 'header',
          section: 0,
          textboxIndex: 0,
          text: 'FIRST\nSECOND\nTHIRD {PAGE}',
        },
      }),
    ).toMatchObject({ committed: true })
    const after = await saved()
    expect(
      after.part.paras.filter((p) => p.textbox).map((p) => p.runs.map((r) => r.text).join('')),
    ).toEqual(['FIRST', 'SECOND', `THIRD ${PAGE_MARK}`])
    expect(after.xml).not.toContain('OLD_HEADER')
    const sections = readSections(after.reopened)
    expect(sections).toHaveLength(2)
    expect(sections[0].headerRefs.default).not.toBe(sections[1].headerRefs.default)
    expect(after.reopened.hfParts![sections[1].headerRefs.default!].text).toContain('OLD_HEADER')
    const xml = new DOMParser().parseFromString(after.xml, 'application/xml')
    for (const box of [...xml.getElementsByTagName('w:txbxContent')]) {
      expect(box.getElementsByTagName('w:p')).toHaveLength(3)
      expect(box.textContent).toContain('THIRD')
    }
    expect([...xml.documentElement.children].filter((p) => p.tagName === 'w:p')).toHaveLength(1)
  } finally {
    runtime.dispose()
  }
})

it('edits a standalone VML textbox through standard MCP and retains its shape', async () => {
  const zip = await JSZip.loadAsync(
    await readFile(resolve(__dirname, 'fixtures/foraoffice-header-textbox.docx')),
  )
  const original = await zip.file('word/header1.xml')!.async('string')
  const fallback = /<mc:Fallback\b[^>]*>([\s\S]*?)<\/mc:Fallback>/.exec(original)![1]
  zip.file(
    'word/header1.xml',
    original.replace(/<mc:AlternateContent>[\s\S]*?<\/mc:AlternateContent>/, fallback),
  )
  const { runtime, saved } = await open(await zip.generateAsync({ type: 'uint8array' }))
  const router = createDocsCommandRouter(runtime)
  const server = createDocsMcpServer({
    version: 'test',
    endpoint: {
      request: (message, signal) => router.handle(message, signal),
      closeCaller: (callerId) => router.disconnect(callerId),
    },
  })
  const client = new Client({ name: 'header-write', version: 'test' }, { capabilities: {} })
  try {
    const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair()
    await server.connect(serverTransport)
    await client.connect(clientTransport)
    const headerTool = (await client.listTools()).tools.find(
      (tool) => tool.name === 'set_header_footer',
    )!
    expect(headerTool.inputSchema).toMatchObject({
      type: 'object',
      required: ['text'],
      anyOf: [{ required: ['kind'] }, { required: ['engineTarget'] }],
      properties: { textboxIndex: { type: 'integer' }, engineTarget: { type: 'object' } },
    })
    const result = await client.callTool({
      name: 'set_header_footer',
      arguments: {
        kind: 'header',
        textboxIndex: 0,
        text: 'VML_HEADER {PAGE} / {NUMPAGES}',
      },
    })
    expect(result).toMatchObject({ structuredContent: { committed: true } })
    const after = await saved()
    expect(after.xml).not.toContain('OLD_HEADER')
    expect(after.xml).toContain('VML_HEADER')
    expect(after.xml).toContain('width:2in;height:2in')
    expect(after.xml).toContain('<v:textbox')
    expect(after.part.text).toContain(PAGE_MARK)
    expect(after.part.text).toContain(TOTAL_PAGES_MARK)
  } finally {
    await client.close()
    await server.close()
    router.dispose()
    runtime.dispose()
  }
})

it.each(['first', 'even'] as const)(
  'refuses structured textboxes before %s display or authored state changes',
  async (view) => {
    const zip = await JSZip.loadAsync(
      await readFile(resolve(__dirname, 'fixtures/foraoffice-header-textbox.docx')),
    )
    zip.file(
      'word/header1.xml',
      (await zip.file('word/header1.xml')!.async('string'))
        .replaceAll('<w:txbxContent>', '<w:txbxContent><w:sdt><w:sdtContent>')
        .replaceAll('</w:txbxContent>', '</w:sdtContent></w:sdt></w:txbxContent>'),
    )
    const body = await zip.file('word/document.xml')!.async('string')
    const reference = /<w:headerReference\b[^>]*w:type="default"[^>]*\/>/.exec(body)![0]
    zip.file(
      'word/document.xml',
      body.replace(
        reference,
        reference +
          reference.replace('"default"', '"first"') +
          reference.replace('"default"', '"even"'),
      ),
    )
    const { caller, runtime, headerFooter, effects } = await open(
      await zip.generateAsync({ type: 'uint8array' }),
    )
    try {
      const before = caller.context()
      const state = effects().state
      const strip = headerFooter
        .read()
        .sections![0].strips.find((s) => s.kind === 'header' && s.view === view)!
      expect(strip.textboxes).toMatchObject([
        { index: 0, editable: false, lifetime: 'current-part' },
      ])
      expect(strip.textboxes![0].text).toContain('OLD_HEADER')
      expect(
        await caller.execute({
          id: 'unsupported',
          name: 'set_header_footer',
          input: {
            kind: 'header',
            view,
            textboxIndex: 0,
            text: 'MUST_NOT_APPLY',
          },
        }),
      ).toMatchObject({ isError: true, committed: false, revision: before.revision })
      expect(effects()).toEqual({ state, firstEnabled: 0, evenEnabled: 0 })
      expect(effects().state).toBe(state)
    } finally {
      runtime.dispose()
    }
  },
)

it('refuses a plain Choice whose compatibility fallback contains protected structure', async () => {
  const zip = await JSZip.loadAsync(
    await readFile(resolve(__dirname, 'fixtures/foraoffice-header-textbox.docx')),
  )
  zip.file(
    'word/header1.xml',
    (await zip.file('word/header1.xml')!.async('string')).replace(
      /<mc:Fallback\b[\s\S]*?<\/mc:Fallback>/,
      (fallback) =>
        fallback
          .replace('<w:txbxContent>', '<w:txbxContent><w:sdt><w:sdtContent>')
          .replace('</w:txbxContent>', '</w:sdtContent></w:sdt></w:txbxContent>'),
    ),
  )
  const { caller, runtime, effects, saved } = await open(
    await zip.generateAsync({ type: 'uint8array' }),
  )
  try {
    const before = effects().state
    const result = await caller.execute({
      id: 'protected-fallback',
      name: 'set_header_footer',
      input: { kind: 'header', text: 'MUST_NOT_APPLY' },
    })
    expect(result).toMatchObject({ isError: true, committed: false })
    expect(effects().state).toBe(before)
    expect((await saved()).xml).toContain('OLD_HEADER')
  } finally {
    runtime.dispose()
  }
})

it.each(['simple', 'complex'])(
  'refuses an unmodelled %s Word field before replacement',
  async (form) => {
    const zip = await JSZip.loadAsync(
      await readFile(resolve(__dirname, 'fixtures/foraoffice-header-textbox.docx')),
    )
    let header = await zip.file('word/header1.xml')!.async('string')
    header =
      form === 'simple'
        ? header.replaceAll(
            '<w:txbxContent>',
            '<w:txbxContent><w:p>' +
              '<w:fldSimple w:instr=" STYLEREF Title "><w:r><w:t>CACHED_TITLE</w:t></w:r></w:fldSimple></w:p>',
          )
        : header.replaceAll(' PAGE ', ' STYLEREF Title ')
    zip.file('word/header1.xml', header)
    const { caller, runtime, headerFooter, effects } = await open(
      await zip.generateAsync({ type: 'uint8array' }),
    )
    try {
      expect(headerFooter.read().sections![0].strips[0].textboxes![0].editable).toBe(false)
      const state = effects().state
      expect(
        await caller.execute({
          id: 'unsupported-field',
          name: 'set_header_footer',
          input: { kind: 'header', text: 'MUST_NOT_APPLY' },
        }),
      ).toMatchObject({ isError: true, committed: false })
      expect(effects().state).toBe(state)
    } finally {
      runtime.dispose()
    }
  },
)

it('refuses run-level content controls within a textbox paragraph', async () => {
  const zip = await JSZip.loadAsync(
    await readFile(resolve(__dirname, 'fixtures/foraoffice-header-textbox.docx')),
  )
  zip.file(
    'word/header1.xml',
    (await zip.file('word/header1.xml')!.async('string')).replaceAll(
      '<w:txbxContent>',
      '<w:txbxContent><w:p><w:sdt><w:sdtPr>' +
        '<w:tag w:val="protected-control"/></w:sdtPr><w:sdtContent>' +
        '<w:r><w:t>CONTROL_TEXT</w:t></w:r></w:sdtContent></w:sdt></w:p>',
    ),
  )
  const { caller, runtime, effects, saved } = await open(
    await zip.generateAsync({ type: 'uint8array' }),
  )
  try {
    const state = effects().state
    expect(
      await caller.execute({
        id: 'run-control',
        name: 'set_header_footer',
        input: { kind: 'header', textboxIndex: 0, text: 'MUST_NOT_APPLY' },
      }),
    ).toMatchObject({ isError: true, committed: false })
    expect(effects().state).toBe(state)
    const xml = new DOMParser().parseFromString((await saved()).xml, 'application/xml')
    expect(xml.getElementsByTagName('w:tag')).toHaveLength(2)
  } finally {
    runtime.dispose()
  }
})

it('refuses whole-part replacement when nested compatibility content cannot map to source textboxes', async () => {
  const zip = await JSZip.loadAsync(
    await readFile(resolve(__dirname, 'fixtures/foraoffice-header-textbox.docx')),
  )
  const header = await zip.file('word/header1.xml')!.async('string')
  const paragraph = header.slice(header.indexOf('<w:p '), header.lastIndexOf('</w:hdr>'))
  const fallback = /<mc:Fallback>([\s\S]*?)<\/mc:Fallback>/.exec(paragraph)![1]
  const nested =
    '<mc:Fallback><mc:AlternateContent><mc:Choice Requires="v">' +
    fallback +
    '</mc:Choice><mc:Fallback>' +
    fallback +
    '</mc:Fallback></mc:AlternateContent>' +
    fallback +
    '</mc:Fallback>'
  zip.file(
    'word/header1.xml',
    header.replace(
      paragraph,
      paragraph.replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/, nested) +
        paragraph.replaceAll('OLD_HEADER', 'TARGET_OLD'),
    ),
  )
  const { caller, runtime, effects, saved } = await open(
    await zip.generateAsync({ type: 'uint8array' }),
  )
  try {
    const state = effects().state
    expect(caller.context().text).toContain('"textEditable":false')
    expect(
      await caller.execute({
        id: 'nested-compatibility',
        name: 'set_header_footer',
        input: { kind: 'header', text: 'MUST_NOT_APPLY' },
      }),
    ).toMatchObject({ isError: true, committed: false })
    expect(effects().state).toBe(state)
    const after = await saved()
    expect(after.xml).toContain('OLD_HEADER')
    expect(after.xml).toContain('TARGET_OLD')
    expect(after.xml).not.toContain('MUST_NOT_APPLY')
  } finally {
    runtime.dispose()
  }
})

it('refuses replacement when an unknown enclosing structure hides a plain source textbox', async () => {
  const zip = await JSZip.loadAsync(
    await readFile(resolve(__dirname, 'fixtures/foraoffice-header-textbox.docx')),
  )
  const header = await zip.file('word/header1.xml')!.async('string')
  const paragraph = header.slice(header.indexOf('<w:p '), header.lastIndexOf('</w:hdr>'))
  zip.file(
    'word/header1.xml',
    header.replace(
      paragraph,
      '<w:customXml w:element="fixture" w:uri="urn:fixture">' + paragraph + '</w:customXml>',
    ),
  )
  const { caller, runtime, effects } = await open(await zip.generateAsync({ type: 'uint8array' }))
  try {
    const state = effects().state
    expect(
      await caller.execute({
        id: 'unmapped-box',
        name: 'set_header_footer',
        input: { kind: 'header', text: 'MUST_NOT_APPLY' },
      }),
    ).toMatchObject({ isError: true, committed: false })
    expect(effects().state).toBe(state)
  } finally {
    runtime.dispose()
  }
})

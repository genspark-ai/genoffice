import { afterEach, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { inlinePictureDrawingXml, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { buildDocContext, serializeRangeToHtml } from '../src/renderer/ai/protocol'
import { executeTool } from '../src/renderer/ai/tools'
import {
  createDocsRuntime,
  type DocsRuntime,
  type DocsRuntimeClient,
} from '../src/renderer/runtime/docs-runtime'
import type { DocsQueryResult } from '../src/shared/document-query'

const numIds = { bullet: null, ordered: null }
const editors: Editor[] = []
const runtimes: DocsRuntime[] = []
afterEach(() => {
  runtimes.splice(0).forEach((runtime) => runtime.dispose())
  editors.splice(0).forEach((editor) => editor.destroy())
})
async function fixture(track?: { author: string }) {
  const parsed = await parseDocx(
    await buildDocx({
      withImage: true,
      bodyXml:
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Historical description</w:t></w:r></w:p>' +
        '<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>' +
        '<w:p><w:fldSimple w:instr="DATE"><w:r><w:t>2020-01-01</w:t></w:r></w:fldSimple></w:p>' +
        `<w:p><w:r>${inlinePictureDrawingXml({ rId: 'rId10', docPrId: 1, widthPx: 100, heightPx: 80 })}</w:r></w:p>` +
        '<w:p><w:r><w:t>Final paragraph</w:t></w:r></w:p>',
    }),
  )
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editors.push(editor)
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  buildDocContext(editor)
  const runtime = createDocsRuntime({
    current: () => ({ editor, identity: editor, numIds, track }),
  })
  runtimes.push(runtime)
  const caller = runtime.connect({ id: 'structural-replacement' })
  caller.context()
  const saved = async () => {
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const zip = await JSZip.loadAsync(await saveDocx(parsed, plan.saveBlocks))
    const xml = await zip.file('word/document.xml')!.async('string')
    const doc = new DOMParser().parseFromString(xml, 'application/xml')
    const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    return {
      sections: doc.getElementsByTagNameNS(ns, 'sectPr').length,
      fields:
        doc.getElementsByTagNameNS(ns, 'fldSimple').length +
        doc.getElementsByTagNameNS(ns, 'instrText').length,
      images: doc.getElementsByTagNameNS(ns, 'drawing').length,
      xml,
    }
  }
  return { editor, saved, caller }
}

async function query(caller: DocsRuntimeClient, start: number, end: number) {
  const result = await caller.execute({
    id: 'structure-query',
    name: 'query_document',
    input: { domain: 'body', startBlockIndex: start, endBlockIndex: end },
  })
  expect(result.isError, result.output).not.toBe(true)
  const value = JSON.parse(result.output) as DocsQueryResult
  expect(value.page.complete).toBe(true)
  value.items.forEach((item) => expect(item.structuresComplete).toBe(true))
  return value
}

it('refuses a mixed read-back rewrite without changing DOCX structures or content', async () => {
  const { editor, saved } = await fixture()
  const before = await saved()
  const html = serializeRangeToHtml(editor, 0, 4).replace(
    'Historical description',
    '{{description}}',
  )
  const result = await executeTool(
    editor,
    {
      id: 'readback',
      name: 'replace_blocks',
      input: { startBlockIndex: 0, endBlockIndex: 4, html },
    },
    numIds,
  )
  expect(result).toMatchObject({ isError: true, mutated: false })
  expect(await saved()).toEqual(before)
})

it('does not pretend that visible field text can round trip its instruction', async () => {
  const { editor, saved } = await fixture()
  const html = serializeRangeToHtml(editor, 2, 2)
  expect(html).toContain('data-doc-readonly="field"')
  const result = await executeTool(
    editor,
    {
      id: 'field',
      name: 'replace_blocks',
      input: { startBlockIndex: 2, endBlockIndex: 2, html },
    },
    numIds,
  )
  expect(result).toMatchObject({ isError: true, mutated: false })
  expect(await saved()).toMatchObject({ sections: 2, fields: 1, images: 1 })
  expect(serializeRangeToHtml(editor, 2, 2, null, (src) => src)).not.toContain('data-doc-readonly')
})

it('keeps deliberate replacement available and reports its structural scope honestly', async () => {
  const { caller, saved } = await fixture()
  const source = await query(caller, 0, 4)
  const result = await caller.execute({
    id: 'replace',
    name: 'replace_blocks',
    input: {
      startBlockIndex: 0,
      endBlockIndex: 4,
      html: '<p>Intentionally rewritten</p>',
      deleteStructures: source.items.flatMap(
        (item) => item.structures?.map((structure) => structure.id) ?? [],
      ),
    },
  })
  expect(result.isError).not.toBe(true)
  expect(result.output).toMatch(/protected.*block/i)
  expect(result.output).toMatch(/field/i)
  expect(result.output).toMatch(/section/i)
  expect(result.output).not.toContain("new blocks kept the replaced blocks' formatting")
  expect(await saved()).toMatchObject({ sections: 1, fields: 0, images: 0 })
})

it('describes pending structural deletion as tracked rather than already removed', async () => {
  const { editor, caller, saved } = await fixture({ author: 'AI' })
  const source = await query(caller, 0, 4)
  const result = await caller.execute({
    id: 'tracked',
    name: 'replace_blocks',
    input: {
      startBlockIndex: 0,
      endBlockIndex: 4,
      html: '<p>Replacement</p>',
      deleteStructures: source.items.flatMap(
        (item) => item.structures?.map((structure) => structure.id) ?? [],
      ),
    },
  })
  expect(result.isError).not.toBe(true)
  expect(result.output).toMatch(/marked for deletion/i)
  expect(editor.state.doc.child(1).attrs.blockRevision.kind).toBe('del')
  expect((await saved()).xml).toMatch(/<w:del[ >]/)
})

it('preserves in-place text editing and explicit section deletion semantics', async () => {
  const { caller, saved } = await fixture()
  const edit = await caller.execute({
    id: 'text',
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
  expect(edit.isError).not.toBe(true)
  expect(await saved()).toMatchObject({ sections: 2, fields: 1, images: 1 })
  const source = await query(caller, 1, 1)
  const deletion = await caller.execute({
    id: 'delete',
    name: 'apply_ops',
    input: {
      ops: [
        {
          op: 'deleteBlocks',
          target: { engine: source.items[0].target },
          deleteStructures: source.items[0].structures?.map((structure) => structure.id),
        },
      ],
    },
  })
  expect(deletion.isError).not.toBe(true)
  expect(await saved()).toMatchObject({ sections: 1, fields: 1, images: 1 })
})

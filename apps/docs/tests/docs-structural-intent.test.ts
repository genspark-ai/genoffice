// Modified for ForaOffice: structural deletion requires exact declared intent.
import { afterEach, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { parseDocx, readSections, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { createDocsRuntime, type DocsRuntime } from '../src/renderer/runtime/docs-runtime'
import { executeTool } from '../src/renderer/ai/tools'

const editors: Editor[] = []
const runtimes: DocsRuntime[] = []
afterEach(() => {
  runtimes.splice(0).forEach((runtime) => runtime.dispose())
  editors.splice(0).forEach((editor) => editor.destroy())
})

async function fixture(bodyXml?: string) {
  const parsed = await parseDocx(
    await buildDocx({
      bodyXml:
        bodyXml ??
        '<w:p><w:pPr><w:mirrorIndents/><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="10000" w:h="15000"/></w:sectPr></w:pPr><w:r><w:t>SECTION END</w:t></w:r></w:p>' +
          '<w:p><w:r><w:t>NEIGHBOR</w:t></w:r></w:p>',
      sectPrXml: '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
    }),
  )
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  const runtime = createDocsRuntime({
    current: () => ({ editor, identity: editor, numIds: { bullet: null, ordered: null } }),
  })
  editors.push(editor)
  runtimes.push(runtime)
  const caller = runtime.connect({ id: 'structure-writer' })
  const savedBytes = async () => {
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    return saveDocx(parsed, plan.saveBlocks)
  }
  const savedXml = async () => {
    const zip = await JSZip.loadAsync(await savedBytes())
    return zip.file('word/document.xml')!.async('string')
  }
  return { editor, caller, savedXml, reopened: async () => parseDocx(await savedBytes()) }
}

it('refuses an ordinary HTML rewrite of a text-bearing section ending before content or revision changes', async () => {
  const { caller, savedXml } = await fixture()
  const before = caller.context()
  const beforeXml = await savedXml()
  const result = await caller.execute({
    id: 'implicit-section',
    name: 'replace_blocks',
    input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>UNACKNOWLEDGED SECTION REWRITE</p>' },
  })
  expect(result).toMatchObject({
    isError: true,
    committed: false,
    mutated: false,
    revision: before.revision,
  })
  expect(caller.context()).toEqual(before)
  expect(await savedXml()).toBe(beforeXml)
})

it('allows the exact section deletion and does not move its properties onto replacement text', async () => {
  const { editor, caller, savedXml, reopened } = await fixture()
  caller.context()
  const result = await executeTool(
    editor,
    {
      id: 'explicit-section',
      name: 'replace_blocks',
      input: {
        startBlockIndex: 0,
        endBlockIndex: 0,
        html: '<p>SECTION REMOVED</p>',
        deleteStructures: ['section:12:13'],
      },
    },
    { bullet: null, ordered: null },
  )
  expect(result.isError).not.toBe(true)
  const doc = new DOMParser().parseFromString(await savedXml(), 'application/xml')
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  expect(doc.getElementsByTagNameNS(ns, 'sectPr')).toHaveLength(1)
  expect(doc.getElementsByTagNameNS(ns, 'p')[0].textContent).toBe('SECTION REMOVED')
  expect(
    doc.getElementsByTagNameNS(ns, 'p')[0].getElementsByTagNameNS(ns, 'mirrorIndents'),
  ).toHaveLength(1)
  expect(doc.getElementsByTagNameNS(ns, 'pgSz')[0].getAttributeNS(ns, 'w')).toBe('12240')
  expect(doc.getElementsByTagNameNS(ns, 'p')[1].textContent).toBe('NEIGHBOR')
  expect(readSections(await reopened())).toHaveLength(1)
})

it('retains the section owner and exact page settings during an in-place text edit', async () => {
  const { caller, savedXml, reopened } = await fixture()
  caller.context()
  const result = await caller.execute({
    id: 'in-place',
    name: 'apply_ops',
    input: {
      ops: [
        {
          op: 'findReplace',
          target: { blockIndexes: [0] },
          find: 'SECTION END',
          replace: 'RENAMED END',
        },
      ],
    },
  })
  expect(result.isError).not.toBe(true)
  const doc = new DOMParser().parseFromString(await savedXml(), 'application/xml')
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  const paragraphs = doc.getElementsByTagNameNS(ns, 'p')
  expect(paragraphs[0].textContent).toBe('RENAMED END')
  expect(paragraphs[0].getElementsByTagNameNS(ns, 'pgSz')[0].getAttributeNS(ns, 'w')).toBe('10000')
  expect(paragraphs[1].getElementsByTagNameNS(ns, 'sectPr')).toHaveLength(0)
  expect(readSections(await reopened())).toHaveLength(2)
})

it('rejects duplicate or unrelated deletion declarations before rewriting a section', async () => {
  const { editor, caller, savedXml } = await fixture()
  const before = caller.context()
  const xml = await savedXml()
  for (const deleteStructures of [
    ['section:12:13', 'section:12:13'],
    ['section:12:13', 'math:1:2'],
  ]) {
    const result = await executeTool(
      editor,
      {
        id: 'mismatch',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>WRONG</p>', deleteStructures },
      },
      { bullet: null, ordered: null },
    )
    expect(result).toMatchObject({ isError: true, mutated: false })
  }
  expect(caller.context()).toEqual(before)
  expect(await savedXml()).toBe(xml)
})

it('refuses ordinary selection replacement of a Word field and leaves its saved instruction active', async () => {
  const { editor, caller, savedXml } = await fixture(
    '<w:p><w:r><w:t xml:space="preserve">LEAD </w:t></w:r><w:fldSimple w:instr="DATE"><w:r><w:t>2020</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> TAIL</w:t></w:r></w:p>',
  )
  editor.commands.setTextSelection({ from: 6, to: 10 })
  const before = caller.context()
  const xml = await savedXml()
  const result = await caller.execute({
    id: 'field-selection',
    name: 'replace_selection',
    input: { html: '2026' },
  })
  expect(result).toMatchObject({
    isError: true,
    committed: false,
    mutated: false,
    revision: before.revision,
  })
  expect(await savedXml()).toBe(xml)
})

it('deletes a field only with its exact intent and does not copy its instruction onto new text', async () => {
  const { editor, caller, savedXml, reopened } = await fixture(
    '<w:p><w:r><w:t xml:space="preserve">LEAD </w:t></w:r><w:fldSimple w:instr="DATE"><w:r><w:t>2020</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> TAIL</w:t></w:r></w:p>',
  )
  editor.commands.setTextSelection({ from: 6, to: 10 })
  caller.context()
  const result = await executeTool(
    editor,
    {
      id: 'delete-field',
      name: 'replace_selection',
      input: { html: '2026', deleteStructures: ['instrField:6:10'] },
    },
    { bullet: null, ordered: null },
  )
  expect(result.isError).not.toBe(true)
  const xml = await savedXml()
  expect(xml).not.toMatch(/<w:(?:fldSimple|instrText)[\s>]/)
  expect((await reopened()).blocks[0].runs?.map((run) => run.text).join('')).toBe('LEAD 2026 TAIL')
})

it('does not mistake reused immutable image nodes for two preserved image identities', async () => {
  const { editor, caller } = await fixture('<w:p><w:r><w:t>IMAGES</w:t></w:r></w:p>')
  const picture = editor.schema.nodes.docInlineImage.create({
    dataUrl: 'data:image/png;base64,AQ==',
  })
  editor.commands.setContent(
    editor.schema.topNodeType.create(null, [
      editor.schema.nodes.docParagraph.create(null, [picture, picture]),
    ]),
  )
  const before = caller.context()
  const result = await caller.execute({
    id: 'one-image-only',
    name: 'replace_blocks',
    input: {
      startBlockIndex: 0,
      endBlockIndex: 0,
      html: `<p><img src="media:${before.documentId}:${before.revision}:0.0"></p>`,
    },
  })
  expect(result).toMatchObject({
    isError: true,
    committed: false,
    mutated: false,
    revision: before.revision,
  })
  expect(caller.context()).toEqual(before)
})

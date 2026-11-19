/**
 * genoffice#2555: a w:ins / w:del whose only run holds an atom (a footnote
 * reference, a line break, …) was invisible to collectRevisions, which scans
 * text nodes only. accept_changes {all:true} then left the wrapper in the
 * file while reporting "No tracked changes remain", and a document whose
 * only revision was such a wrapper was rejected outright. The fixtures are
 * the issue's synthetic documents (footnotes part + reference wired in).
 */
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { afterEach, describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { applyRevisions, collectRevisions } from '../src/renderer/editor/revisions'

const REV = 'w:author="A" w:date="2026-01-01T00:00:00Z"'
const OTHER = `<w:ins w:id="9" ${REV}><w:r><w:t> other change</w:t></w:r></w:ins>`
const FN_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Note.</w:t></w:r></w:p></w:footnote></w:footnotes>'
const FN_REL =
  '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'

const CASES: Record<string, string> = {
  fnref_ins: `<w:p><w:r><w:t>Text.</w:t></w:r><w:ins w:id="1" ${REV}><w:r><w:footnoteReference w:id="1"/></w:r></w:ins>${OTHER}</w:p>`,
  fnref_del: `<w:p><w:r><w:t>Text.</w:t></w:r><w:del w:id="1" ${REV}><w:r><w:footnoteReference w:id="1"/></w:r></w:del>${OTHER}</w:p>`,
  br_ins: `<w:p><w:r><w:t>A</w:t></w:r><w:ins w:id="1" ${REV}><w:r><w:br/></w:r></w:ins>${OTHER}</w:p>`,
  text_and_fnref_one_ins: `<w:p><w:ins w:id="1" ${REV}><w:r><w:t>new </w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:ins></w:p>`,
  fnref_ins_only: `<w:p><w:r><w:t>Text.</w:t></w:r><w:ins w:id="1" ${REV}><w:r><w:footnoteReference w:id="1"/></w:r></w:ins></w:p>`,
}

const liveEditors: Editor[] = []
afterEach(() => {
  for (const editor of liveEditors.splice(0)) editor.destroy()
})

async function buildCase(bodyXml: string): Promise<Uint8Array> {
  return buildDocx({
    bodyXml,
    extraRels: FN_REL,
    extraParts: [
      {
        path: 'word/footnotes.xml',
        xml: FN_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
      },
    ],
  })
}

async function bodyXmlOf(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes)
  return (await zip.file('word/document.xml')!.async('string')) ?? ''
}

const pendingWrappers = (xml: string) => (xml.match(/<w:(?:ins|del) [^>]*w:id="/g) ?? []).length

for (const [name, bodyXml] of Object.entries(CASES)) {
  it(`${name}: the atomic run is a listed revision`, async () => {
    const parsed = await parseDocx(await buildCase(bodyXml))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    liveEditors.push(editor)
    expect(collectRevisions(editor.state.doc).length).toBeGreaterThan(0)
  })

  it(`${name}: accept clears the wrapper`, async () => {
    const parsed = await parseDocx(await buildCase(bodyXml))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    liveEditors.push(editor)
    applyRevisions(editor, collectRevisions(editor.state.doc), 'accept')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xml = await bodyXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(pendingWrappers(xml)).toBe(0)
  })
}

describe('accept / reject the atomic run itself', () => {
  it('accepting an inserted footnote reference keeps the marker, unwrapped', async () => {
    const parsed = await parseDocx(await buildCase(CASES.fnref_ins!))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    liveEditors.push(editor)
    applyRevisions(editor, collectRevisions(editor.state.doc), 'accept')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xml = await bodyXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(xml).toContain('<w:footnoteReference w:id="1"/>')
    expect(pendingWrappers(xml)).toBe(0)
  })

  it('accepting a deleted footnote reference removes the marker; rejecting keeps it', async () => {
    for (const mode of ['accept', 'reject'] as const) {
      const parsed = await parseDocx(await buildCase(CASES.fnref_del!))
      const editor = new Editor({
        element: document.createElement('div'),
        extensions: editorExtensions,
        content: blocksToPmDoc(parsed.blocks) as never,
      })
      liveEditors.push(editor)
      applyRevisions(editor, collectRevisions(editor.state.doc), mode)
      const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
      const xml = await bodyXmlOf(await saveDocx(parsed, plan.saveBlocks))
      const refs = (xml.match(/<w:footnoteReference/g) ?? []).length
      if (mode === 'accept') expect(refs).toBe(0)
      else expect(refs).toBe(1)
      expect(pendingWrappers(xml)).toBe(0)
    }
  })

  it('accepting an inserted line break keeps the break, unwrapped', async () => {
    const parsed = await parseDocx(await buildCase(CASES.br_ins!))
    const editor = new Editor({
      element: document.createElement('div'),
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks) as never,
    })
    liveEditors.push(editor)
    applyRevisions(editor, collectRevisions(editor.state.doc), 'accept')
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xml = await bodyXmlOf(await saveDocx(parsed, plan.saveBlocks))
    expect(xml).toContain('<w:br/>')
    expect(pendingWrappers(xml)).toBe(0)
  })
})

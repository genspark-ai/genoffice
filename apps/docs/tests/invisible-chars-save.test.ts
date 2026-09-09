import { Editor } from '@tiptap/core'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { executeOps } from '../src/renderer/ai/ops'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { editorExtensions } from '../src/renderer/editor/extensions'

// Word stores U+00A0 as a literal character inside w:t (no xml:space needed).
// The field NodeView used to re-read its cached result from the DOM on every
// save and strip every NBSP, so an untouched citation field lost the character
// and the shortened text was redistributed over the original w:t slots.
const FIELD_TEXT = 'Accepted in Radiology Experimental'
const FIELD_PARA =
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {} </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  `<w:r><w:rPr><w:i/></w:rPr><w:t>${FIELD_TEXT}</w:t></w:r>` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

const INVISIBLE = 'a b​c d'
const BODY_PARA =
  `<w:p><w:r><w:t>${INVISIBLE}</w:t></w:r>` +
  '<w:r><w:t>e</w:t><w:noBreakHyphen/><w:t>f</w:t><w:softHyphen/><w:t>g</w:t></w:r></w:p>'

async function openDoc(bodyXml: string) {
  const source = await buildDocx({ bodyXml })
  const parsed = await parseDocx(source)
  const host = document.createElement('div')
  document.body.appendChild(host)
  const editor = new Editor({
    element: host,
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks) as never,
  })
  return { source, parsed, editor, host }
}

async function documentXml(bytes: Uint8Array): Promise<string> {
  return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string')
}

describe('invisible characters survive saving', () => {
  it('keeps an untouched field result byte-exact across the pre-save commit', async () => {
    const { source, parsed, editor, host } = await openDoc(FIELD_PARA)
    window.dispatchEvent(new Event('ai-docs-commit-tables'))
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(0)
    expect(await saveDocx(parsed, plan.saveBlocks)).toEqual(source)
    editor.destroy()
    host.remove()
  })

  it('keeps the NBSP when the field result is edited in place', async () => {
    const { parsed, editor, host } = await openDoc(FIELD_PARA)
    const text = host.querySelector<HTMLElement>('.doc-field-text')!
    expect(text.textContent).toBe(FIELD_TEXT)
    text.textContent = FIELD_TEXT + '!'
    text.dispatchEvent(new Event('input', { bubbles: true }))
    window.dispatchEvent(new Event('ai-docs-commit-tables'))
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    expect(plan.changedCount).toBe(1)
    const xml = await documentXml(await saveDocx(parsed, plan.saveBlocks))
    expect(xml).toContain(`<w:t>${FIELD_TEXT}!</w:t>`)
    expect(xml).toContain('ZOTERO_ITEM')
    editor.destroy()
    host.remove()
  })

  it('round-trips NBSP/ZWSP/NNBSP and hyphen elements through a regenerated paragraph', async () => {
    const { parsed, editor, host } = await openDoc(BODY_PARA)
    const outcome = executeOps(editor, [{ op: 'findReplace', find: INVISIBLE, replace: INVISIBLE }])
    expect(outcome.ok).toBe(true)
    expect(outcome.results[0]).toMatchObject({ matched: 1 })
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const xml = await documentXml(await saveDocx(parsed, plan.saveBlocks))
    expect(xml).toContain(INVISIBLE)
    expect(xml).toMatch(
      />e<\/w:t><w:noBreakHyphen\/><w:t[^>]*>f<\/w:t><w:softHyphen\/><w:t[^>]*>g</,
    )
    expect(xml).not.toMatch(/[‑­]/)
    editor.destroy()
    host.remove()
  })
})

import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import JSZip from 'jszip'
import { mergeStyleXml, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { executeOps } from '../src/renderer/ai/ops'
import { resolveStyleDefinition } from '../src/renderer/ai/style-ops'
import { withStyleUpsert } from '../src/renderer/editor/design-history'

const editors = new Set<Editor>()
afterEach(() => {
  for (const e of editors) e.destroy()
  editors.clear()
})

const PLAIN_P = '<w:p><w:pPr><w:keepNext/></w:pPr><w:r><w:t>Body</w:t></w:r></w:p>'
const EXACT_P =
  '<w:p><w:pPr><w:keepNext/><w:spacing w:line="280" w:lineRule="exact"/></w:pPr>' +
  '<w:r><w:t>Body</w:t></w:r></w:p>'

async function openEditor(bodyXml: string) {
  const parsed = await parseDocx(await buildDocx({ bodyXml }))
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
  })
  editors.add(editor)
  editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
  return { editor, parsed }
}

async function saveXml(editor: Editor, parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
  const saved = await saveDocx(parsed, plan.saveBlocks)
  const zip = await JSZip.loadAsync(saved)
  return zip.file('word/document.xml')!.async('string')
}

const format = (lineSpacing: unknown) => ({
  op: 'setParagraphFormat',
  target: { blockIndexes: [0] },
  lineSpacing,
})

describe('setParagraphFormat exact line spacing', () => {
  it('a length string writes an exact rule that survives save and reopen', async () => {
    const { editor, parsed } = await openEditor(PLAIN_P)
    expect(executeOps(editor, [format('30pt')]).ok).toBe(true)
    const xml = await saveXml(editor, parsed)
    expect(xml).toContain('<w:spacing w:line="600" w:lineRule="exact"/>')
    expect(xml).toContain('<w:keepNext/>')
    const reopened = await parseDocx(await buildDocx({ bodyXml: xml.match(/<w:p>.*<\/w:p>/)![0] }))
    expect(reopened.blocks[0]).toMatchObject({
      format: { lineRule: 'exact', lineRawTwips: 600 },
    })
  })

  it('a multiple replaces an existing exact rule instead of being shadowed by it', async () => {
    const { editor, parsed } = await openEditor(EXACT_P)
    expect(executeOps(editor, [format(1.5)]).ok).toBe(true)
    const xml = await saveXml(editor, parsed)
    expect(xml).toContain('<w:spacing w:line="360" w:lineRule="auto"/>')
    expect(xml).not.toContain('w:lineRule="exact"')
  })

  it('null clears the exact rule and its twips', async () => {
    const { editor, parsed } = await openEditor(EXACT_P)
    expect(executeOps(editor, [format(null)]).ok).toBe(true)
    const xml = await saveXml(editor, parsed)
    expect(xml).not.toContain('w:line=')
    expect(xml).not.toContain('w:lineRule')
    expect(xml).toContain('<w:keepNext/>')
  })

  it.each(['0pt', '-12pt', '30', 'thirty', '30 lines', '99999in'])(
    'rejects the malformed or non-positive length %j',
    async (value) => {
      const { editor } = await openEditor(PLAIN_P)
      const out = executeOps(editor, [format(value)])
      expect(out.ok).toBe(false)
      expect(out.error).toContain('lineSpacing')
    },
  )
})

describe('define_style exact line spacing', () => {
  const resolve = (lineSpacing: unknown) =>
    resolveStyleDefinition({ styleId: 'Body', paragraph: { lineSpacing } }, [])

  it('a length string becomes an exact style rule', () => {
    const r = resolve('30pt')
    if ('error' in r) throw new Error(r.error)
    expect(r.value.upsert.pPr).toEqual({ lineFixed: { rule: 'exact', twips: 600 } })
    const existing =
      '<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/>' +
      '<w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:style>'
    expect(mergeStyleXml(existing, r.value.upsert)).toContain(
      '<w:spacing w:after="120" w:line="600" w:lineRule="exact"/>',
    )
  })

  it('a multiple still writes an auto rule', () => {
    const r = resolve(1.5)
    if ('error' in r) throw new Error(r.error)
    expect(r.value.upsert.pPr).toEqual({ lineSpacing: 1.5 })
  })

  it.each([0, -1, '0pt', '30', 'thirty'])('rejects %j', (value) => {
    const r = resolve(value)
    expect('error' in r && r.error).toContain('paragraph.lineSpacing')
  })

  it.each([
    [1.5, 'w:line="360" w:lineRule="auto"'],
    [null, '<w:qFormat/></w:style>'],
  ])('%j after an unsaved exact rule replaces it', (next, expected) => {
    let pending = {}
    for (const value of ['30pt', next]) {
      const r = resolve(value)
      if ('error' in r) throw new Error(r.error)
      pending = withStyleUpsert(pending, r.value.upsert)
    }
    const xml = mergeStyleXml(null, (pending as Record<string, never>).Body)
    expect(xml).toContain(expected)
    expect(xml).not.toContain('exact')
  })
})

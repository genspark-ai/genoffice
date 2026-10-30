import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { Editor } from '@tiptap/core'
import {
  countTrackedChangesXml,
  parseDocx,
  saveDocx,
  type HeaderFooter,
  type HfPartInfo,
  type NoteInfo,
  type SectionInfo,
} from '@genoffice/docx-engine'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { buildReviewEntries, filterRevisions } from '../src/renderer/editor/revision-filters'
import {
  STORY_LABEL,
  collectHfRevisions,
  collectNoteRevisions,
  collectTextboxRevisions,
  liveStoryRevision,
  ownHfSlots,
  resolveHfValue,
  resolveNotes,
  resolveTextboxes,
  resolutionOf,
} from '../src/renderer/editor/story-revisions'
import { listRevisionEntries } from '../src/renderer/ai/revision-ops'
import type { HfSectionState } from '../src/renderer/hf-sections'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'

const ANN = { author: 'Ann', date: '2026-09-29T12:23:00Z', id: '7' }
const BOB = { author: 'Bob', date: '2026-09-30T08:00:00Z', id: '8' }
const REV_ANN = ' w:id="7" w:author="Ann" w:date="2026-09-29T12:23:00Z"'
const REV_BOB = ' w:id="8" w:author="Bob" w:date="2026-09-30T08:00:00Z"'
const NS = ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

const FOOTNOTES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:footnotes${NS}>` +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Kept </w:t></w:r>' +
  `<w:del${REV_ANN}><w:r><w:delText>gone</w:delText></w:r></w:del>` +
  `<w:ins${REV_BOB}><w:r><w:t>added</w:t></w:r></w:ins></w:p></w:footnote>` +
  '</w:footnotes>'

const TXBX_PARAGRAPH =
  '<w:p><w:r><w:drawing><wp:anchor behindDoc="0" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:extent cx="6038850" cy="1704975"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6038850" cy="1704975"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  '<wps:txbx><w:txbxContent><w:p><w:r><w:t xml:space="preserve">Box </w:t></w:r>' +
  `<w:del${REV_ANN}><w:r><w:delText>old</w:delText></w:r></w:del>` +
  `<w:ins${REV_BOB}><w:r><w:t>new</w:t></w:r></w:ins></w:p></w:txbxContent></wps:txbx></wps:wsp>` +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>'

const noteDoc = () =>
  buildDocx({
    bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>',
    extraRels:
      '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
    extraParts: [
      {
        path: 'word/footnotes.xml',
        xml: FOOTNOTES_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
      },
    ],
  })

const section = (lastBlockIndex: number, headerRefs: SectionInfo['headerRefs']): SectionInfo => ({
  settings: {} as SectionInfo['settings'],
  startType: 'nextPage',
  firstBlockIndex: 0,
  lastBlockIndex,
  sectPrXml: '',
  titlePg: false,
  headerRefs,
  footerRefs: {},
})

const headerPart = (): HfPartInfo => ({
  text: 'Kept added',
  hasPageNumber: false,
  paras: [
    {
      runs: [{ text: 'Kept ' }, { text: 'gone', del: ANN }, { text: 'added', ins: BOB }],
    },
    {
      runs: [{ text: 'Bold', bold: true, rPrChange: { ...ANN, old: { bold: false } } }],
    },
  ],
})

/** two sections sharing one header part, one pending footer edit */
const hfState = (): HfSectionState => ({
  sections: [section(10, { default: 'rH' }), section(20, { default: 'rH' })],
  edits: {},
  links: {},
  parsed: { hfParts: { rH: headerPart() } },
})

describe('enumeration', () => {
  it('footnote changes list once per contiguous record, anchored at the note reference', async () => {
    const parsed = await parseDocx(await noteDoc())
    const editor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    const revs = collectNoteRevisions(editor.state.doc, 'footnote', parsed.footnotes)
    expect(revs.map((r) => [r.kind, r.author, r.story.text, r.story.kind])).toEqual([
      ['ins', 'Bob', 'added', 'footnote'],
      ['del', 'Ann', 'gone', 'footnote'],
    ])
    let refPos = -1
    editor.state.doc.descendants((node, pos) => {
      if (node.type.name === 'docNoteRef') refPos = pos
    })
    expect(refPos).toBeGreaterThan(0)
    expect(revs.every((r) => r.from === refPos)).toBe(true)
    const entries = buildReviewEntries(editor.state.doc, revs)
    expect(entries.map((e) => e.excerpt)).toEqual(['added', 'gone'])
    expect(STORY_LABEL[revs[0].story.kind]).toBe('appRevLocFootnote')
    editor.destroy()
  })

  it('a header part shared by two sections lists its changes once, formatting included', () => {
    const revs = collectHfRevisions(hfState())
    expect(revs.map((r) => [r.kind, r.story.text, r.story.para])).toEqual([
      ['ins', 'added', 0],
      ['del', 'gone', 0],
      ['rPrChange', 'Bold', 1],
    ])
    expect(revs[0].story.hf).toEqual({ section: 0, kind: 'header', variant: 'default' })
    expect(filterRevisions(revs, { insDel: true, format: false, hiddenAuthors: [] })).toHaveLength(
      2,
    )
    const editor = new Editor({
      extensions: editorExtensions,
      content: { type: 'doc', content: [] },
    })
    expect(listRevisionEntries(editor.state.doc, revs).map((e) => e.blockIndex)).toEqual([
      -1, -1, -1,
    ])
    editor.destroy()
    expect(
      filterRevisions(revs, { insDel: true, format: true, hiddenAuthors: ['Ann'] }),
    ).toHaveLength(1)
  })

  it('textbox changes carry the anchor position; the AI list labels the location', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: TXBX_PARAGRAPH }))
    const editor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    const revs = collectTextboxRevisions(editor.state.doc)
    expect(revs.map((r) => [r.kind, r.story.text])).toEqual([
      ['ins', 'new'],
      ['del', 'old'],
    ])
    expect(revs[0].story.textbox).toEqual({ pos: 0, box: 0 })
    const entries = listRevisionEntries(editor.state.doc, revs)
    expect(entries.map((e) => [e.id, e.location, e.text])).toEqual([
      ['r1', 'textbox', 'new'],
      ['r2', 'textbox', 'old'],
    ])
    editor.destroy()
  })
})

describe('accept / reject', () => {
  it('a footnote deletion accepted leaves the rich runs, the text and the saved part clean', async () => {
    const parsed = await parseDocx(await noteDoc())
    const editor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    const revs = collectNoteRevisions(editor.state.doc, 'footnote', parsed.footnotes)
    const del = revs.find((r) => r.kind === 'del')!
    const notes = resolveNotes(parsed.footnotes, [del.story], 'accept')
    expect(notes[0].richParas![0].map((r) => r.text)).toEqual(['Kept ', 'added'])
    expect(notes[0].text).toBe('Kept added')
    expect(notes[0].revisionsResolved).toEqual([{ mode: 'accept', filter: { ids: ['7'] } }])
    // the stale entry still resolves after the models moved on
    const left = collectNoteRevisions(editor.state.doc, 'footnote', notes)
    expect(left.map((r) => r.kind)).toEqual(['ins'])
    expect(liveStoryRevision(left, del)).toBeUndefined()
    expect(
      liveStoryRevision(
        left,
        revs.find((r) => r.kind === 'ins')!,
      ),
    ).toBeDefined()

    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks, { footnotes: notes })
    const xml = await (await JSZip.loadAsync(saved)).file('word/footnotes.xml')!.async('string')
    expect(xml).not.toContain('gone')
    expect(xml).toMatch(/<w:ins w:id="8"/)
    expect(countTrackedChangesXml(xml)).toBe(1)

    const rejected = resolveNotes(notes, [liveStoryRevision(left, revs[0])!.story], 'reject')
    expect(rejected[0].text).toBe('Kept ')
    expect(rejected[0].revisionsResolved).toHaveLength(2)
    const again = await saveDocx(parsed, plan.saveBlocks, { footnotes: rejected })
    const xml2 = await (await JSZip.loadAsync(again)).file('word/footnotes.xml')!.async('string')
    expect(countTrackedChangesXml(xml2)).toBe(0)
    expect(xml2).not.toContain('added')
    editor.destroy()
  })

  it('a header change on a parsed part keeps the part bytes; a pending edit regenerates', () => {
    const state = hfState()
    const revs = collectHfRevisions(state)
    const value = state.parsed!.hfParts!.rH as HeaderFooter
    const accepted = resolveHfValue(value, [revs[1].story], 'accept', false)
    expect(accepted.keepContent).toBe(true)
    expect(accepted.paras![0].runs.map((r) => r.text)).toEqual(['Kept ', 'added'])
    expect(accepted.text).toBe('Kept addedBold')
    expect(accepted.revisionsResolved).toEqual([{ mode: 'accept', filter: { ids: ['7'] } }])

    const formatBack = resolveHfValue(accepted, [revs[2].story], 'reject', true)
    expect(formatBack.keepContent).toBe(true)
    expect(formatBack.paras![1].runs[0]).toEqual({ text: 'Bold' })
    expect(formatBack.revisionsResolved).toHaveLength(2)

    const typed: HeaderFooter = { ...value, paras: value.paras }
    const all = resolveHfValue(
      typed,
      revs.map((r) => r.story),
      'reject',
      true,
      true,
    )
    expect(all.keepContent).toBeUndefined()
    expect(all.paras![0].runs.map((r) => r.text)).toEqual(['Kept ', 'gone'])
    expect(all.revisionsResolved).toEqual([{ mode: 'reject' }])
    expect(collectHfRevisions({ ...state, edits: { '20:header': all } })).toHaveLength(0)
  })

  it('a section unlinked from a shared part lists its own copy separately', () => {
    const state = hfState()
    // Link to Previous off: section 2 keeps a copy of the inherited paragraphs as its own part
    const unlinked: HfSectionState = {
      ...state,
      sections: [state.sections[0], { ...state.sections[1], headerRefs: {} }],
      edits: { '20:header': { text: 'Kept added', paras: headerPart().paras } },
    }
    const slots = ownHfSlots(unlinked)
    expect(slots.map((s) => s.si)).toEqual([0, 1])
    expect(collectHfRevisions(unlinked)).toHaveLength(6)
  })

  it('a textbox change rewritten into the node saves without markup', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: TXBX_PARAGRAPH }))
    const editor = new Editor({
      extensions: editorExtensions,
      content: blocksToPmDoc(parsed.blocks),
    })
    const revs = collectTextboxRevisions(editor.state.doc)
    const node = editor.state.doc.nodeAt(0)!
    const boxes = resolveTextboxes(
      node.attrs.textboxes,
      revs.map((r) => r.story),
      'reject',
    )
    editor.view.dispatch(
      editor.state.tr.setNodeMarkup(0, undefined, { ...node.attrs, textboxes: boxes }),
    )
    expect(collectTextboxRevisions(editor.state.doc)).toHaveLength(0)
    const plan = pmDocToSavePlan(editor.getJSON() as PmNode, parsed.blocks)
    const saved = await saveDocx(parsed, plan.saveBlocks)
    const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
    expect(countTrackedChangesXml(xml)).toBe(0)
    expect(xml).toContain('old')
    expect(xml).not.toContain('new')
    const reparsed = await parseDocx(saved)
    expect(reparsed.blocks[0].textboxes![0].paras[0].runs.map((r) => r.text)).toEqual(['Box old'])
    editor.destroy()
  })

  it('records without a w:id resolve by author and date', () => {
    const note: NoteInfo = {
      id: '1',
      text: 'x',
      richParas: [[{ text: 'x', ins: { author: 'Zed', date: '2026-01-01T00:00:00Z' } }]],
    }
    expect(resolutionOf(collectNoteRevisionsOf(note), 'accept')).toEqual([
      { mode: 'accept', filter: { author: 'Zed', date: '2026-01-01T00:00:00Z' } },
    ])
  })
})

function collectNoteRevisionsOf(note: NoteInfo) {
  const editor = new Editor({ extensions: editorExtensions, content: { type: 'doc', content: [] } })
  const out = collectNoteRevisions(editor.state.doc, 'footnote', [note]).map((r) => r.story)
  editor.destroy()
  return out
}

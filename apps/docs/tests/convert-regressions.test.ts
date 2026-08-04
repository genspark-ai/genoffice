import { describe, expect, it } from 'vitest'
import type { Block, GeneratedBlock } from '@genoffice/docx-engine'
import {
  blocksToPmDoc,
  inlineToRuns,
  pmDocToSavePlan,
  signatureOfBlock,
  signatureOfGenerated,
  type PmNode,
} from '../src/renderer/editor/convert'

describe('inlineToRuns hard break after atomic runs', () => {
  it('does not append the break char to a footnote reference run', () => {
    const inline: PmNode[] = [
      { type: 'text', text: 'before' },
      { type: 'docNoteRef', attrs: { kind: 'footnote', id: '3', num: 3 } },
      { type: 'hardBreak' },
      { type: 'text', text: 'after' },
    ]
    const runs = inlineToRuns(inline)
    const refRun = runs.find((r) => r.noteRef)
    expect(refRun?.text).toBe('3')
    expect(runs.map((r) => r.text).join('')).toBe('before3\nafter')
  })

  it('does not append the break char to an inline math run', () => {
    const inline: PmNode[] = [
      { type: 'docInlineMath', attrs: { omml: '<m:oMath/>', text: 'x' } },
      { type: 'hardBreak', attrs: { pageBreak: true } },
    ]
    const runs = inlineToRuns(inline)
    expect(runs[0]).toEqual({ text: 'x', math: { omml: '<m:oMath/>' } })
    expect(runs[1]).toEqual({ text: '\f' })
  })
})

describe('paragraph format signatures', () => {
  const base: Block = {
    id: 'b0',
    type: 'paragraph',
    docxIndex: 0,
    originalXml: '<w:p/>',
    runs: [{ text: 'a' }],
  }

  it('detects lineRule/lineRawTwips differences', () => {
    const generated: GeneratedBlock = { type: 'paragraph', runs: [{ text: 'a' }] }
    const withRule = signatureOfBlock({
      ...base,
      format: { lineRule: 'exact', lineRawTwips: 480 },
    })
    const without = signatureOfGenerated(generated)
    expect(withRule).not.toBe(without)
    const matching = signatureOfGenerated({
      ...generated,
      format: { lineRule: 'exact', lineRawTwips: 480 },
    })
    expect(withRule).toBe(matching)
  })

  it('detects bidi differences', () => {
    const ltr = signatureOfBlock(base)
    const rtl = signatureOfBlock({ ...base, format: { bidi: true } })
    expect(ltr).not.toBe(rtl)
  })
})

describe('pasted copy of an anchored protected block', () => {
  it('materializes a duplicated image from its preview bytes instead of dropping it', () => {
    const block: Block = {
      id: 'b0',
      type: 'image',
      docxIndex: 0,
      originalXml: '<w:p><w:r><w:drawing/></w:r></w:p>',
      imageDataUrl: 'data:image/png;base64,QUJD',
      imageWidthPx: 120,
      imageHeightPx: 80,
    }
    const doc = blocksToPmDoc([block])
    doc.content!.push(doc.content![0])
    const plan = pmDocToSavePlan(doc, [block])
    expect(plan.saveBlocks).toHaveLength(2)
    expect(plan.saveBlocks[0]).toEqual({ kind: 'original', docxIndex: 0 })
    expect(plan.saveBlocks[1]).toEqual({
      kind: 'image',
      image: { base64: 'QUJD', mime: 'image/png', widthPx: 120, heightPx: 80 },
    })
  })

  it('clones originalXml for a duplicated passthrough, stripping bookmark/comment anchors', () => {
    const block: Block = {
      id: 'b0',
      type: 'passthrough',
      docxIndex: 0,
      originalXml:
        '<w:p><w:bookmarkStart w:id="1" w:name="bm"/><w:r><w:t>x</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>',
    }
    const doc = blocksToPmDoc([block])
    doc.content!.push(doc.content![0])
    const plan = pmDocToSavePlan(doc, [block])
    expect(plan.saveBlocks[0]).toEqual({ kind: 'original', docxIndex: 0 })
    const copy = plan.saveBlocks[1]
    if (copy.kind !== 'xml') throw new Error(`expected xml, got ${copy.kind}`)
    expect(copy.xml).toBe('<w:p><w:r><w:t>x</w:t></w:r></w:p>')
  })
})

describe('split twin bookmark/comment anchors', () => {
  it('emits bookmarks and comment endpoints only on the anchor, not the twin', () => {
    const block: Block = {
      id: 'b0',
      type: 'paragraph',
      docxIndex: 0,
      originalXml: '<w:p/>',
      runs: [{ text: 'hello world' }],
      bookmarks: ['bm'],
      hiddenBookmarks: ['_Toc1'],
      commentStarts: ['7'],
      commentEnds: ['7'],
    }
    const node = blocksToPmDoc([block]).content![0]
    const doc: PmNode = {
      type: 'doc',
      content: [
        { ...node, content: [{ type: 'text', text: 'hello ' }] },
        { ...node, content: [{ type: 'text', text: 'world' }] },
      ],
    }
    const plan = pmDocToSavePlan(doc, [block])
    const [anchor, twin] = plan.saveBlocks
    if (anchor.kind !== 'generated' || twin.kind !== 'generated') {
      throw new Error('expected two generated blocks')
    }
    expect(anchor.block.bookmarks).toEqual(['bm'])
    expect(anchor.block.commentEnds).toEqual(['7'])
    expect(twin.block.bookmarks).toBeUndefined()
    expect(twin.block.hiddenBookmarks).toBeUndefined()
    expect(twin.block.commentStarts).toBeUndefined()
    expect(twin.block.commentEnds).toBeUndefined()
  })
})

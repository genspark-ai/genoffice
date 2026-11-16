// Modified for ForaOffice: keep unsaved section boundaries tied to their live editor nodes.
import { describe, expect, it } from 'vitest'
import { Schema } from '@tiptap/pm/model'
import {
  liveSections,
  liveSectionsForDocument,
  sameSectionInfos,
  sourceSectionIndex,
} from '../src/renderer/pagination-sections'
import { sectionIndexAtBlock } from '../src/renderer/ai/pending-sections'

type SectionInfo = Parameters<typeof liveSections>[0][number]
type BlockBox = Parameters<typeof liveSections>[1][number]

const section = (first: number, last: number): SectionInfo =>
  ({
    firstBlockIndex: first,
    lastBlockIndex: last,
    startType: 'nextPage',
  }) as unknown as SectionInfo
const present = (...docxIndexes: number[]): BlockBox[] =>
  docxIndexes.map((docxIndex) => ({ docxIndex }) as unknown as BlockBox)

describe('sameSectionInfos', () => {
  const sections = [section(0, 9), section(10, 19), section(20, 29)]

  it('treats the merged list a streaming open recomputes per chunk as unchanged', () => {
    // the second section's break block has not streamed in yet: it merges into the third
    const a = liveSections(sections, present(9, 29))
    const b = liveSections(sections, present(9, 29))
    expect(a).not.toBe(b)
    expect(a).not.toBe(sections)
    expect(sameSectionInfos(a, b)).toBe(true)
  })

  it('reports a change once another break block lands', () => {
    const merged = liveSections(sections, present(9, 29))
    const complete = liveSections(sections, present(9, 19, 29))
    expect(complete).toBe(sections)
    expect(sameSectionInfos(merged, complete)).toBe(false)
  })

  it('is identity for the same array and false for a different length', () => {
    expect(sameSectionInfos(sections, sections)).toBe(true)
    expect(sameSectionInfos(sections, sections.slice(1))).toBe(false)
  })
})

describe('live pending section boundaries', () => {
  const schema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      text: { group: 'inline' },
      docParagraph: { group: 'block', content: 'text*', attrs: { docxIndex: { default: null } } },
      docProtected: {
        group: 'block',
        atom: true,
        attrs: {
          docxIndex: { default: null },
          genXml: { default: null },
          pendingSectionId: { default: null },
          blockRevision: { default: null },
        },
      },
    },
  })
  const paragraph = (index: number) =>
    schema.node('docParagraph', { docxIndex: index }, schema.text('Body'))
  const pending = (id: string, deleted = false) =>
    schema.node('docProtected', {
      pendingSectionId: id,
      genXml: '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>',
      blockRevision: deleted ? { kind: 'del' } : null,
    })
  const pendingInfo = (id: string, first: number, last: number) =>
    Object.assign(section(first, last), {
      pendingBreak: true as const,
      pendingBreakId: id,
    })

  it('keeps an empty first section and routes its closing break separately from following content', () => {
    const first = pendingInfo('first', 0, -1)
    const last = section(0, 1)
    const doc = schema.node('doc', null, [pending('first'), paragraph(0), paragraph(1)])
    const live = liveSectionsForDocument([first, last], doc)
    expect(live).toEqual([first, last])
    expect(sectionIndexAtBlock(doc, live, 0)).toBe(0)
    expect(sectionIndexAtBlock(doc, live, 1)).toBe(1)
  })

  it('removes only the missing pending boundary when two sections share an original block index, and restores it on undo', () => {
    const first = pendingInfo('first', 0, 0)
    const second = pendingInfo('second', 1, 0)
    const last = section(1, 1)
    const sections = [first, second, last]
    const removed = schema.node('doc', null, [paragraph(0), pending('second'), paragraph(1)])
    expect(liveSectionsForDocument(sections, removed)).toEqual([
      { ...second, firstBlockIndex: 0 },
      last,
    ])
    const restored = schema.node('doc', null, [
      paragraph(0),
      pending('first'),
      pending('second'),
      paragraph(1),
    ])
    expect(liveSectionsForDocument(sections, restored)).toEqual(sections)
    expect(sections).toEqual([first, second, last])
  })

  it('does not count a tracked-deleted generated break as a live section', () => {
    const first = pendingInfo('deleted', 0, 0)
    const last = section(1, 1)
    const doc = schema.node('doc', null, [paragraph(0), pending('deleted', true), paragraph(1)])
    expect(liveSectionsForDocument([first, last], doc)).toEqual([{ ...last, firstBlockIndex: 0 }])
  })

  it('maps a merged live section and its cursor back to the correct authoritative entry', () => {
    const sections = [section(0, 1), section(2, 3), section(4, 5)]
    const doc = schema.node('doc', null, [
      paragraph(0),
      paragraph(2),
      paragraph(3),
      paragraph(4),
      paragraph(5),
    ])
    const live = liveSectionsForDocument(sections, doc)
    expect(live).toHaveLength(2)
    expect(sourceSectionIndex(live, sections[1])).toBe(0)
    expect(sourceSectionIndex(sections, live[0])).toBe(1)
    const owner = live[sectionIndexAtBlock(doc, live, 0)]
    expect(sourceSectionIndex(sections, owner)).toBe(1)
  })

  it('distinguishes empty sections with equal source boundaries and never substitutes an original for a missing pending identity', () => {
    const first = pendingInfo('first', 0, 0)
    const second = pendingInfo('second', 1, 0)
    expect(sourceSectionIndex([first, second], second)).toBe(1)
    expect(sourceSectionIndex([first, second], { ...second })).toBe(1)
    expect(sourceSectionIndex([section(0, 0)], second)).toBe(-1)
  })
})

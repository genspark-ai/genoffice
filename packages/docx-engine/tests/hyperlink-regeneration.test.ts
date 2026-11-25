import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'
import { packageXmlProblems } from './helpers/strict-xml'

it('keeps adjacent links with the same target and different history settings separate', async () => {
  const target = 'https://example.com/doc.html#section2'
  const bytes = await buildDocx({
    bodyXml:
      '<w:p><w:hyperlink r:id="rId11" w:history="0"><w:r><w:t>first</w:t></w:r></w:hyperlink>' +
      '<w:hyperlink r:id="rId11"><w:r><w:t>second</w:t></w:r></w:hyperlink></w:p>',
    extraRels:
      '<Relationship Id="rId11" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ' +
      `Target="${target}" TargetMode="External"/>`,
  })
  const parsed = await parseDocx(bytes)
  const paragraph = parsed.blocks[0]
  if (paragraph.type !== 'paragraph' || !paragraph.runs) {
    throw new Error('Expected a paragraph with runs')
  }
  const output = await saveDocx(parsed, [
    {
      kind: 'generated',
      docxIndex: paragraph.docxIndex ?? undefined,
      block: { type: 'paragraph', runs: paragraph.runs },
    },
  ])
  const zip = await JSZip.loadAsync(output)
  const xml = await zip.file('word/document.xml')!.async('string')
  expect(xml.match(/<w:hyperlink\b[^>]*>/g)).toEqual([
    '<w:hyperlink r:id="rId11" w:history="0">',
    '<w:hyperlink r:id="rId11">',
  ])

  const reopened = (await parseDocx(output)).blocks[0]
  if (reopened.type !== 'paragraph' || !reopened.runs) {
    throw new Error('Expected a paragraph with runs')
  }
  expect(reopened.runs.map((run) => run.link?.href)).toEqual([target, target])
})

const locations = [
  {
    name: 'external',
    external: true,
    attrs: ' w:tgtFrame="_blank" w:docLocation="page-two"',
    frame: '_blank',
    location: 'page-two',
  },
  { name: 'external defaults', external: true, attrs: '', frame: undefined, location: undefined },
  {
    name: 'explicit empty',
    external: true,
    attrs: ' w:tgtFrame="" w:docLocation=""',
    frame: '',
    location: '',
  },
  {
    name: 'internal bookmark',
    external: false,
    attrs: ' w:tgtFrame="_self" w:docLocation="page-two"',
    frame: '_self',
    location: 'page-two',
  },
  { name: 'internal defaults', external: false, attrs: '', frame: undefined, location: undefined },
  {
    name: 'escaped attributes',
    external: true,
    attrs: ' w:tgtFrame="frame&amp;&quot;&lt;$$" w:docLocation="page&amp;&quot;&lt;$$"',
    frame: 'frame&"<$$',
    location: 'page&"<$$',
  },
]

it.each(locations)(
  'preserves $name link location when editing adjacent ordinary text',
  async (test) => {
    const untouched = '<w:p><w:r><w:t>untouched</w:t></w:r></w:p>'
    const bytes = await buildDocx({
      bodyXml:
        `<w:p><w:hyperlink${test.external ? ' r:id="rId11"' : ''} w:anchor="ch2"${test.attrs}><w:r><w:t>link</w:t></w:r></w:hyperlink><w:r><w:t>after</w:t></w:r></w:p>` +
        untouched,
      extraRels: test.external
        ? '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/doc" TargetMode="External"/>'
        : '',
    })
    const parsed = await parseDocx(bytes)
    const runs = parsed.blocks[0].runs!.map((run) =>
      run.text === 'after' ? { ...run, text: 'end' } : run,
    )
    const saved = await saveDocx(parsed, [
      { kind: 'generated', docxIndex: 0, block: { type: 'paragraph', runs } },
      { kind: 'original', docxIndex: 1 },
    ])
    expect(await packageXmlProblems(saved)).toEqual([])
    const out = await JSZip.loadAsync(saved)
    const xml = await out.file('word/document.xml')!.async('string')
    const open = xml.match(/<w:hyperlink\b[^>]*>/)![0]
    expect(open).toContain('w:anchor="ch2"')
    const escape = (value: string) =>
      value.replaceAll('&', '&amp;').replaceAll('"', '&quot;').replaceAll('<', '&lt;')
    for (const [attr, value] of [
      ['tgtFrame', test.frame],
      ['docLocation', test.location],
    ] as const) {
      if (value === undefined) expect(open).not.toContain(`w:${attr}=`)
      else expect(open).toContain(`w:${attr}="${escape(value)}"`)
    }
    const reopened = await parseDocx(saved)
    expect(reopened.blocks[0].runs!.map((run) => run.text).join('')).toBe('linkend')
    expect(reopened.blocks[0].runs!.at(-1)?.link).toBeUndefined()
    expect(reopened.blocks[0].runs![0].link).toMatchObject({
      href: test.external ? 'https://example.org/doc#ch2' : '#ch2',
    })
    const again = await saveDocx(reopened, [
      {
        kind: 'generated',
        docxIndex: 0,
        block: { type: 'paragraph', runs: reopened.blocks[0].runs! },
      },
      { kind: 'original', docxIndex: 1 },
    ])
    const againZip = await JSZip.loadAsync(again)
    expect(
      (await againZip.file('word/document.xml')!.async('string')).match(
        /<w:hyperlink\b[^>]*>/,
      )?.[0],
    ).toBe(open)
    expect(xml).toContain(untouched)
    const original = await JSZip.loadAsync(bytes)
    for (const [path, entry] of Object.entries(original.files)) {
      if (entry.dir || path === 'word/document.xml') continue
      expect(await out.file(path)!.async('uint8array'), path).toEqual(
        await entry.async('uint8array'),
      )
    }
  },
)

it('keeps adjacent links with different target frames and document locations separate', async () => {
  const attrs = [
    'w:tgtFrame="_blank" w:docLocation="one"',
    'w:tgtFrame="_self" w:docLocation="one"',
    'w:tgtFrame="_self" w:docLocation="two"',
  ]
  const bytes = await buildDocx({
    bodyXml:
      '<w:p>' +
      attrs
        .map(
          (attr, i) =>
            `<w:hyperlink r:id="rId11" ${attr}><w:r><w:t>link${i}</w:t></w:r></w:hyperlink>`,
        )
        .join('') +
      '</w:p>',
    extraRels:
      '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/doc" TargetMode="External"/>',
  })
  const parsed = await parseDocx(bytes)
  expect(parsed.blocks[0].runs).toHaveLength(3)
  const saved = await saveDocx(parsed, [
    { kind: 'generated', docxIndex: 0, block: { type: 'paragraph', runs: parsed.blocks[0].runs! } },
  ])
  const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')
  expect(xml.match(/<w:hyperlink\b[^>]*>/g)).toEqual(
    attrs.map((attr) => `<w:hyperlink r:id="rId11" ${attr}>`),
  )
})

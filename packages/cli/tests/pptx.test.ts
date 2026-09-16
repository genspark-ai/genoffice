import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { inlineLocalFiles } from '../src/formats/pptx'
import { run, tempDir } from './helpers'

const INCH = 914400

function opsFile(dir: string, name: string, ops: unknown[]): string {
  const path = join(dir, name)
  writeFileSync(path, JSON.stringify(ops))
  return path
}

describe('genoffice create/slides (pptx ops)', () => {
  it('builds a deck from ops, reads it back, and edits it in place', async () => {
    const dir = tempDir()
    const create = opsFile(dir, 'create.json', [
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'textbox',
        offset: { x: INCH, y: INCH, cx: 8 * INCH, cy: INCH },
        paragraphs: [{ runs: [{ text: 'Hello genoffice', bold: true, fontSize: 32 }] }],
      },
      { op: 'addBlankSlide', target: { slide: 0 } },
      {
        op: 'addTable',
        target: { slide: 1 },
        rows: 2,
        cols: 3,
        offset: { x: INCH, y: 2 * INCH, cx: 6 * INCH, cy: 2 * INCH },
      },
      { op: 'setNotes', target: { slide: 1 }, text: 'speaker notes' },
    ])
    const out = join(dir, 'deck.pptx')
    const created = await run(['create', '--type', 'pptx', '--ops', create, '--out', out, '--json'])
    expect(created.code).toBe(0)
    expect(created.json().detail).toMatchObject({ applied: true, ops: 4 })
    const zip = await JSZip.loadAsync(readFileSync(out))
    expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).toContain('Hello genoffice')
    expect(zip.file('ppt/slides/slide2.xml')).not.toBeNull()

    const read = await run(['slides', 'read', out, '--json'])
    expect(read.code).toBe(0)
    const deck = read.json().detail
    expect(deck.slides).toBe(2)
    expect(deck.pages[0].id).toMatch(/^s_/)
    const styled = deck.pages[0].elements.find(
      (e: { text?: string }) => e.text === 'Hello genoffice',
    )
    expect(styled.effective).toMatchObject({ fontSizePt: 32, bold: true, italic: false })
    expect(styled.effective.src.fontSize).toBe('run')
    expect(styled.effective.src.bold).toBe('run')
    expect(styled.effective.src.italic).not.toBe('run')
    expect(typeof styled.effective.src.fontFamily).toBe('string')
    const textbox = deck.pages[0].elements.find(
      (e: { text?: string }) => e.text === 'Hello genoffice',
    )
    expect(textbox.id).toMatch(/^e_/)
    expect(textbox.box).toEqual({ x: INCH, y: INCH, cx: 8 * INCH, cy: INCH })
    const table = deck.pages[1].elements.find((e: { type: string }) => e.type === 'table')
    expect(table).toMatchObject({ rows: 2, cols: 3 })

    const edit = opsFile(dir, 'edit.json', [
      {
        op: 'setText',
        target: { slide: 0, el: textbox.id },
        paragraphs: [{ runs: [{ text: 'Edited by genoffice' }] }],
      },
    ])
    const applied = await run(['slides', 'apply', out, '--ops', edit, '--json'])
    expect(applied.code).toBe(0)
    expect(applied.json().detail.records[0]).toMatchObject({ op: 'setText', slide: 's_1' })
    const after = await run(['slides', 'read', out, '--slide', '0', '--json'])
    expect(after.json().detail.pages).toHaveLength(1)
    expect(after.json().detail.pages[0].elements[0].text).toBe('Edited by genoffice')
  })

  it('read --full keeps whole text, every table row and the speaker notes', async () => {
    const dir = tempDir()
    const long = 'word '.repeat(100).trim()
    const create = opsFile(dir, 'create.json', [
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'textbox',
        offset: { x: INCH, y: INCH, cx: 8 * INCH, cy: INCH },
        paragraphs: [{ runs: [{ text: long }] }],
      },
      {
        op: 'addTable',
        target: { slide: 0 },
        rows: 5,
        cols: 2,
        offset: { x: INCH, y: 3 * INCH, cx: 6 * INCH, cy: 2 * INCH },
      },
      { op: 'setNotes', target: { slide: 0 }, text: 'read me aloud' },
    ])
    const out = join(dir, 'deck.pptx')
    expect((await run(['create', '--type', 'pptx', '--ops', create, '--out', out])).code).toBe(0)

    const preview = (await run(['slides', 'read', out, '--json'])).json().detail.pages[0]
    expect(preview.elements[0].text).toBe(`${long.slice(0, 300)}…(+${long.length - 300} chars)`)
    expect(preview.elements[0].truncated).toBe(true)
    expect(preview.elements[1].text.split('\n')).toHaveLength(3)
    expect(preview.elements[1].truncated).toBe(true)
    expect(preview.notes).toBeUndefined()
    const clipped = (await run(['slides', 'read', out, '--max-chars', '40', '--json'])).json()
      .detail.pages[0]
    expect(clipped.elements[0].text).toBe(`${long.slice(0, 40)}…(+${long.length - 40} chars)`)

    const full = (await run(['slides', 'read', out, '--full', '--json'])).json().detail.pages[0]
    expect(full.elements[0].text).toBe(long)
    expect(full.elements[1].text.split('\n')).toHaveLength(5)
    expect(full.elements[1].truncated).toBeUndefined()
    expect(full.notes).toBe('read me aloud')
  })

  it('returns guided op errors as a usage failure and leaves the file untouched', async () => {
    const dir = tempDir()
    const out = join(dir, 'deck.pptx')
    await run([
      'create',
      '--type',
      'pptx',
      '--ops',
      opsFile(dir, 'c.json', [{ op: 'addBlankSlide', target: { slide: 0 } }]),
      '--out',
      out,
    ])
    const before = readFileSync(out)
    const bad = opsFile(dir, 'bad.json', [
      { op: 'addElement', target: { slide: 0 }, kind: 'textbox' },
    ])
    const r = await run(['slides', 'apply', out, '--ops', bad, '--json'])
    expect(r.code).toBe(1)
    expect(r.json().detail.failures[0]).toMatchObject({ index: 0, op: 'addElement' })
    expect(r.json().detail.failures[0].error).toMatch(/offset|Usage/)
    expect(readFileSync(out).equals(before)).toBe(true)

    const dry = await run([
      'slides',
      'apply',
      out,
      '--ops',
      opsFile(dir, 'ok.json', [{ op: 'addBlankSlide', target: { slide: 0 } }]),
      '--dry-run',
      '--json',
    ])
    expect(dry.code).toBe(0)
    expect(dry.json().detail.plan).toHaveLength(1)
    expect(readFileSync(out).equals(before)).toBe(true)

    const notJson = opsFile(dir, 'x.json', [])
    writeFileSync(notJson, '{not json')
    expect(
      (await run(['create', '--type', 'pptx', '--ops', notJson, '--out', join(dir, 'y.pptx')]))
        .code,
    ).toBe(1)
    expect(
      (await run(['create', '--type', 'docx', '--ops', bad, '--out', join(dir, 'z.docx')])).code,
    ).toBe(1)
  })

  it('prints the op guides that ship with the ops', async () => {
    const catalog = await run(['guide', 'slides'])
    expect(catalog.code).toBe(0)
    expect(catalog.stdout).toContain('insert')
    const insert = await run(['guide', 'slides', 'insert'])
    expect(insert.stdout).toContain('addElement')
    const index = await run(['guide', 'slides', '--index'])
    expect(index.stdout).toMatch(/^## text/m)
    expect((await run(['guide', 'slides', 'nope'])).code).toBe(1)
    expect((await run(['guide', 'pdf'])).code).toBe(1)
  })

  it('inlines local files into nested byte fields and creates --out directories', async () => {
    const dir = tempDir()
    const img = join(dir, 'pic.png')
    writeFileSync(img, Buffer.from('89504e470d0a1a0a', 'hex'))
    const [fill, media, pic, other] = inlineLocalFiles(
      [
        { op: 'setImageFill', target: { slide: 0, el: 'e_1' }, source: { bytes: 'pic.png' } },
        { op: 'addMedia', target: { slide: 0 }, bytes: img, poster: { bytes: 'pic.png' } },
        { op: 'addPicture', target: { slide: 0 }, bytes: 'data:image/png;base64,AAAA' },
        { op: 'setText', target: { slide: 0, el: 'e_1' }, paragraphs: [] },
      ],
      { cwd: dir, env: {} },
    )
    expect((fill!.source as { bytes: unknown }).bytes).toBeInstanceOf(Uint8Array)
    expect(media!.bytes).toBeInstanceOf(Uint8Array)
    expect((media!.poster as { bytes: unknown }).bytes).toBeInstanceOf(Uint8Array)
    expect(pic!.bytes).toBe('data:image/png;base64,AAAA')
    expect(other!.paragraphs).toEqual([])

    const out = join(dir, 'deck.pptx')
    await run([
      'create',
      '--type',
      'pptx',
      '--ops',
      opsFile(dir, 'c.json', [{ op: 'addBlankSlide', target: { slide: 0 } }]),
      '--out',
      out,
    ])
    const nested = join(dir, 'a/b/copy.pptx')
    const r = await run([
      'slides',
      'apply',
      out,
      '--ops',
      opsFile(dir, 'e.json', [{ op: 'addBlankSlide', target: { slide: 0 } }]),
      '--out',
      nested,
      '--json',
    ])
    expect(r.code).toBe(0)
    expect(existsSync(nested)).toBe(true)
  })
})

describe('local picture files', () => {
  it('fills the ext the op needs from the file name', async () => {
    const { inlineLocalFiles } = await import('../src/formats/pptx')
    const dir = tempDir()
    writeFileSync(join(dir, 'pic.PNG'), Buffer.from('89504e470d0a1a0a', 'hex'))
    const [op] = inlineLocalFiles(
      [{ op: 'addPicture', target: { slide: 0 }, bytes: 'pic.PNG' } as never],
      { cwd: dir, env: {} },
    )
    expect((op as { ext?: string }).ext).toBe('png')
    expect((op as unknown as { bytes: unknown }).bytes).toBeInstanceOf(Uint8Array)
  })
})

describe('slides apply runs ops in order', () => {
  async function deck(dir: string): Promise<string> {
    const ops = join(dir, 'create.json')
    writeFileSync(
      ops,
      JSON.stringify([
        {
          op: 'addElement',
          target: { slide: 0 },
          kind: 'textbox',
          offset: { x: 914400, y: 685800, cx: 7315200, cy: 914400 },
          paragraphs: [{ runs: [{ text: 'Hello' }] }],
        },
      ]),
    )
    const out = join(dir, 'deck.pptx')
    expect((await run(['create', '--type', 'pptx', '--ops', ops, '--out', out])).code).toBe(0)
    return out
  }
  const textbox = (slide: number, text: string) => ({
    op: 'addElement',
    target: { slide },
    kind: 'textbox',
    offset: { x: 914400, y: 914400, cx: 3657600, cy: 914400 },
    paragraphs: [{ runs: [{ text }] }],
  })

  it('lets a later op target the slide an earlier op added, atomically', async () => {
    const dir = tempDir()
    const pptx = await deck(dir)
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([{ op: 'addBlankSlide', target: { slide: 0 } }, textbox(1, 'Budget')]),
    )
    const r = await run(['slides', 'apply', pptx, '--ops', ops, '--json'])
    expect(r.code).toBe(0)
    expect(r.json().detail.ops).toBe(2)
    const read = await run(['slides', 'read', pptx, '--slide', '1', '--json'])
    expect(JSON.stringify(read.json().detail.pages[0])).toContain('Budget')
  })

  it('writes nothing when a later op fails in atomic mode, keeps the rest in per_op mode', async () => {
    const dir = tempDir()
    const pptx = await deck(dir)
    const before = readFileSync(pptx)
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([{ op: 'addBlankSlide', target: { slide: 0 } }, textbox(7, 'nope')]),
    )
    const atomic = await run(['slides', 'apply', pptx, '--ops', ops, '--json'])
    expect(atomic.code).toBe(1)
    expect(atomic.json().message).toMatch(/^op 1 \(addElement\) rejected/)
    expect(readFileSync(pptx).equals(before)).toBe(true)
    const perOp = await run([
      'slides',
      'apply',
      pptx,
      '--ops',
      ops,
      '--isolation',
      'per_op',
      '--json',
    ])
    expect(perOp.code).toBe(0)
    expect(perOp.json().detail.failures).toHaveLength(1)
    expect(perOp.json().detail.failures[0].index).toBe(1)
    expect(perOp.json().detail.failures[0].error).not.toMatch(/atomic|whole transaction/)
    expect((await run(['info', pptx, '--json'])).json().detail.slides).toBe(2)
  })

  it('dry-run validates the ordered batch without writing', async () => {
    const dir = tempDir()
    const pptx = await deck(dir)
    const before = readFileSync(pptx)
    const ops = join(dir, 'ops.json')
    writeFileSync(
      ops,
      JSON.stringify([{ op: 'addBlankSlide', target: { slide: 0 } }, textbox(1, 'Budget')]),
    )
    const r = await run(['slides', 'apply', pptx, '--ops', ops, '--dry-run', '--json'])
    expect(r.code).toBe(0)
    expect(r.json().summary).toContain('2 of 2 ops validated')
    expect(r.json().detail.applied).toBe(false)
    expect(r.json().detail.plan).toEqual(['0: addBlankSlide on s_1', '1: addElement on s_2'])
    const mixed = join(dir, 'mixed.json')
    writeFileSync(
      mixed,
      JSON.stringify([
        { op: 'addBlankSlide', target: { slide: 0 } },
        textbox(7, 'nope'),
        textbox(1, 'ok'),
      ]),
    )
    const perOp = await run([
      'slides',
      'apply',
      pptx,
      '--ops',
      mixed,
      '--dry-run',
      '--isolation',
      'per_op',
      '--json',
    ])
    expect(perOp.code).toBe(0)
    expect(perOp.json().detail.plan.map((l: string) => l.split(':')[0])).toEqual(['0', '2'])
    expect(perOp.json().detail.failures[0].index).toBe(1)
    expect(readFileSync(pptx).equals(before)).toBe(true)
  })
})

describe('genoffice slides (arrangement, table styles, layouts)', () => {
  it('lists layouts, adds a slide by layout name, styles a table and glues a connector', async () => {
    const dir = tempDir()
    const out = join(dir, 'arrange.pptx')
    const create = opsFile(dir, 'create.json', [
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'rect',
        offset: { x: 0, y: 0, cx: INCH, cy: INCH },
      },
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'ellipse',
        offset: { x: 4 * INCH, y: 2 * INCH, cx: INCH, cy: INCH },
      },
      {
        op: 'addTable',
        target: { slide: 0 },
        rows: 2,
        cols: 2,
        offset: { x: 0, y: 4 * INCH, cx: 4 * INCH, cy: INCH },
      },
    ])
    expect(
      (await run(['create', '--type', 'pptx', '--ops', create, '--out', out, '--json'])).code,
    ).toBe(0)

    const read = (await run(['slides', 'read', out, '--layouts', '--json'])).json().detail
    expect(read.layouts).toEqual([{ index: 0, name: 'Blank', type: 'blank', placeholders: [] }])
    const [rect, ellipse, table] = read.pages[0].elements.map((e: { id: string }) => e.id)
    expect(read.pages[0].elements[2].style.name).toBe('Medium Style 2 - Accent 1')

    const edit = opsFile(dir, 'edit.json', [
      { op: 'alignElements', target: { slide: 0 }, els: [rect, ellipse], mode: 'top' },
      { op: 'addConnector', target: { slide: 0 }, from: rect, to: ellipse, kind: 'straight' },
      {
        op: 'setTableStyle',
        target: { slide: 0, el: table },
        styleId: 'Light Style 1 - Accent 2',
        firstRow: true,
        bandRow: false,
        lastRow: true,
      },
      { op: 'addSlideWithLayout', layout: 'Blank' },
    ])
    const applied = await run(['slides', 'apply', out, '--ops', edit, '--json'])
    expect(applied.code).toBe(0)
    expect(applied.json().detail).toMatchObject({ applied: true, ops: 4 })

    const after = (await run(['slides', 'read', out, '--json'])).json().detail
    expect(after.slides).toBe(2)
    const els = after.pages[0].elements
    expect(els.find((e: { id: string }) => e.id === ellipse).box.y).toBe(0)
    expect(els.find((e: { id: string }) => e.id === table).style).toEqual({
      id: '{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}',
      name: 'Light Style 1 - Accent 2',
      flags: ['firstRow', 'lastRow'],
    })
    const zip = await JSZip.loadAsync(readFileSync(out))
    const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    expect(xml).toMatch(/<a:stCxn id="\d+" idx="3"\/><a:endCxn id="\d+" idx="2"\/>/)
    expect(zip.file('ppt/slides/slide2.xml')).not.toBeNull()

    const bad = await run([
      'slides',
      'apply',
      out,
      '--ops',
      opsFile(dir, 'bad.json', [{ op: 'addSlideWithLayout', layout: 'Title Only' }]),
      '--json',
    ])
    expect(bad.code).toBe(1)
    expect(bad.json().message).toContain('0: "Blank"')
  })
})

describe('slides animations and equations', () => {
  it('adds, lists and removes animations and writes an a14:m equation', async () => {
    const dir = tempDir()
    const out = join(dir, 'anim.pptx')
    const create = opsFile(dir, 'create.json', [
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'textbox',
        offset: { x: INCH, y: INCH, cx: 4 * INCH, cy: INCH },
        paragraphs: [{ runs: [{ text: 'Hello' }] }],
      },
    ])
    expect(
      (await run(['create', '--type', 'pptx', '--ops', create, '--out', out, '--json'])).code,
    ).toBe(0)
    const first = JSON.parse((await run(['slides', 'read', out, '--json'])).stdout)
    const el = first.detail.pages[0].elements[0].id
    expect(first.detail.pages[0].animations).toBeUndefined()

    const edit = opsFile(dir, 'edit.json', [
      {
        op: 'addAnimation',
        target: { slide: 0, el },
        effect: 'flyIn',
        direction: 'left',
        duration: 800,
      },
      { op: 'addAnimation', target: { slide: 0, el }, effect: 'fadeOut', trigger: 'afterPrev' },
      { op: 'insertEquation', target: { slide: 0, el }, latex: '\\frac{a}{b}' },
    ])
    const applied = await run(['slides', 'apply', out, '--ops', edit, '--json'])
    expect(applied.code, applied.stderr).toBe(0)

    const read = JSON.parse((await run(['slides', 'read', out, '--slide', '0', '--json'])).stdout)
    const page = read.detail.pages[0]
    expect(
      page.animations.map((a: { seq: number; effect: string; kind: string; trigger: string }) => [
        a.seq,
        a.effect,
        a.kind,
        a.trigger,
      ]),
    ).toEqual([
      [0, 'flyIn', 'entrance', 'onClick'],
      [1, 'fadeOut', 'exit', 'afterPrev'],
    ])
    expect(page.animations[0].el).toBe(el)
    expect(page.animations[0].direction).toBe('left')
    expect(page.animations[0].durationMs).toBe(800)
    expect(page.elements[0].text).toContain('a/b')

    const zip = await JSZip.loadAsync(readFileSync(out))
    const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    expect(slideXml).toContain('<p:timing>')
    expect(slideXml).toContain('presetClass="entr" presetSubtype="8"')
    expect(slideXml).toContain('<a14:m><m:oMathPara')
    expect(slideXml).toMatch(/<mc:Fallback><a:r>[\s\S]*<a:t>a\/b<\/a:t>/)

    const remove = opsFile(dir, 'remove.json', [
      { op: 'removeAnimation', target: { slide: 0 }, seq: 0 },
    ])
    expect((await run(['slides', 'apply', out, '--ops', remove, '--json'])).code).toBe(0)
    const after = JSON.parse((await run(['slides', 'read', out, '--slide', '0', '--json'])).stdout)
    expect(after.detail.pages[0].animations.map((a: { effect: string }) => a.effect)).toEqual([
      'fadeOut',
    ])

    const bad = await run([
      'slides',
      'apply',
      out,
      '--ops',
      opsFile(dir, 'bad.json', [{ op: 'reorderAnimation', target: { slide: 0 }, seq: 4, to: 0 }]),
      '--json',
    ])
    expect(bad.code).toBe(1)
    expect(JSON.parse(bad.stdout).message).toContain('0-0')
  })
})

describe('setLink / setText named show actions', () => {
  it('writes ppaction://hlinkshowjump links, reads them back, rejects unknown actions', async () => {
    const dir = tempDir()
    const out = join(dir, 'actions.pptx')
    const create = opsFile(dir, 'create.json', [
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'rect',
        offset: { x: INCH, y: INCH, cx: 2 * INCH, cy: INCH },
      },
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'textbox',
        offset: { x: INCH, y: 3 * INCH, cx: 6 * INCH, cy: INCH },
        paragraphs: [{ runs: [{ text: 'Back to start' }] }],
      },
    ])
    expect((await run(['create', '--type', 'pptx', '--ops', create, '--out', out])).code).toBe(0)
    const before = await run(['slides', 'read', out, '--json'])
    const [rect, box] = before.json().detail.pages[0].elements as Array<{ id: string }>

    const edit = opsFile(dir, 'edit.json', [
      {
        op: 'setLink',
        target: { slide: 0, el: rect!.id },
        link: { kind: 'action', action: 'nextslide' },
      },
      {
        op: 'setText',
        target: { slide: 0, el: box!.id },
        paragraphs: [
          { runs: [{ text: 'Back to start', link: { kind: 'action', action: 'firstslide' } }] },
        ],
      },
    ])
    const applied = await run(['slides', 'apply', out, '--ops', edit, '--json'])
    expect(applied.code).toBe(0)

    const zip = await JSZip.loadAsync(readFileSync(out))
    const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string')
    expect(slideXml).toContain('r:id="" action="ppaction://hlinkshowjump?jump=nextslide"')
    expect(slideXml).toContain('r:id="" action="ppaction://hlinkshowjump?jump=firstslide"')
    const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string')
    expect(rels).not.toContain('hyperlink')

    const after = await run(['slides', 'read', out, '--json'])
    const [rect2] = after.json().detail.pages[0].elements as Array<{ link?: unknown }>
    expect(rect2!.link).toEqual({ kind: 'action', action: 'nextslide' })

    const bad = opsFile(dir, 'bad.json', [
      {
        op: 'setLink',
        target: { slide: 0, el: rect!.id },
        link: { kind: 'action', action: 'home' },
      },
    ])
    const rejected = await run(['slides', 'apply', out, '--ops', bad, '--json'])
    expect(rejected.code).not.toBe(0)
    expect(JSON.stringify(rejected.json())).toContain('nextslide, previousslide, firstslide')
  })
})

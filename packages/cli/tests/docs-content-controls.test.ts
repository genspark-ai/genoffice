import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { run, tempDir } from './helpers'

/**
 * Block-level content controls (w:sdt): a law firm wraps the parts of a
 * filing that must not be reflowed in one, with a tag and a lock, and finds
 * them again by tag (genoffice#2062).
 */
const SDT =
  '<w:sdt><w:sdtPr><w:alias w:val="Caption"/><w:tag w:val="example:caption"/>' +
  '<w:id w:val="4242"/><w:lock w:val="sdtLocked"/><w:richText/></w:sdtPr>' +
  '<w:sdtContent>' +
  '<w:p><w:r><w:t>Caption line one</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>Caption line two</w:t></w:r></w:p>' +
  '</w:sdtContent></w:sdt>'

const BODY =
  '<w:p><w:r><w:t>Body paragraph before.</w:t></w:r></w:p>' +
  SDT +
  '<w:p><w:r><w:t>Body paragraph after.</w:t></w:r></w:p>'

const plainDocx = async (): Promise<Uint8Array> =>
  buildDocx({
    bodyXml:
      '<w:p><w:r><w:t>Body paragraph before.</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>Caption line one</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>Caption line two</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t>Body paragraph after.</w:t></w:r></w:p>',
  })

const documentXml = async (path: string): Promise<string> => {
  const zip = await JSZip.loadAsync(readFileSync(path))
  return zip.file('word/document.xml')!.async('string')
}

it('reads the control a block sits in', async () => {
  const dir = tempDir()
  const input = join(dir, 'in.docx')
  writeFileSync(input, await buildDocx({ bodyXml: BODY }))

  const result = await run(['docs', 'read', input, '--json'])
  expect(result.code, result.stderr).toBe(0)
  const blocks = JSON.parse(result.stdout).detail.items as Array<{
    index: number
    text: string
    contentControl?: { tag: string; title: string; lock: string; id: string }
  }>

  const captioned = blocks.filter((b) => b.contentControl)
  expect(captioned.map((b) => b.text)).toEqual(['Caption line one', 'Caption line two'])
  expect(captioned[0]!.contentControl).toEqual({
    tag: 'example:caption',
    title: 'Caption',
    lock: 'sdtLocked',
    id: '4242',
  })
  // the paragraphs outside the control carry no control
  expect(blocks.filter((b) => !b.contentControl).map((b) => b.text)).toEqual([
    'Body paragraph before.',
    'Body paragraph after.',
  ])
})

it('wraps a block range in one content control', async () => {
  const dir = tempDir()
  const input = join(dir, 'in.docx')
  const ops = join(dir, 'ops.json')
  writeFileSync(input, await plainDocx())
  writeFileSync(
    ops,
    JSON.stringify([
      {
        op: 'wrap_content_control',
        startBlockIndex: 1,
        endBlockIndex: 2,
        tag: 'example:caption',
        title: 'Caption',
        lock: 'sdtLocked',
      },
    ]),
  )

  const result = await run(['docs', 'apply', input, '--ops', ops, '--json'])
  expect(result.code, result.stderr).toBe(0)

  const xml = await documentXml(input)
  expect(xml).toContain('<w:sdt>')
  expect(xml).toContain('w:val="example:caption"')
  expect(xml).toContain('w:val="sdtLocked"')
  expect(xml).toContain('w:val="Caption"')
  // the wrapped blocks are inside the control, and the others are not
  const sdt = /<w:sdt>[\s\S]*?<\/w:sdt>/.exec(xml)?.[0] ?? ''
  expect(sdt).toContain('Caption line one')
  expect(sdt).toContain('Caption line two')
  expect(sdt).not.toContain('Body paragraph before.')
  expect(sdt).not.toContain('Body paragraph after.')

  // and the saved document reads back with the control on both blocks
  const read = await run(['docs', 'read', input, '--json'])
  const blocks = JSON.parse(read.stdout).detail.items as Array<{
    text: string
    contentControl?: { tag: string; lock: string }
  }>
  const captioned = blocks.filter((b) => b.contentControl)
  expect(captioned.map((b) => b.text)).toEqual(['Caption line one', 'Caption line two'])
  expect(captioned[0]!.contentControl!.tag).toBe('example:caption')
  expect(captioned[0]!.contentControl!.lock).toBe('sdtLocked')
})

it('unwraps the control with a tag', async () => {
  const dir = tempDir()
  const input = join(dir, 'in.docx')
  const ops = join(dir, 'ops.json')
  writeFileSync(input, await buildDocx({ bodyXml: BODY }))
  writeFileSync(ops, JSON.stringify([{ op: 'unwrap_content_control', tag: 'example:caption' }]))

  const result = await run(['docs', 'apply', input, '--ops', ops, '--json'])
  expect(result.code, result.stderr).toBe(0)

  const xml = await documentXml(input)
  expect(xml).not.toContain('<w:sdt>')
  expect(xml).not.toContain('example:caption')
  // the text is still there, as plain body paragraphs
  expect(xml).toContain('Caption line one')
  expect(xml).toContain('Caption line two')

  const read = await run(['docs', 'read', input, '--json'])
  const blocks = JSON.parse(read.stdout).detail.items as Array<{
    text: string
    contentControl?: unknown
  }>
  expect(blocks.filter((b) => b.contentControl)).toEqual([])
})

it('rejects a wrap without a usable block range', async () => {
  const dir = tempDir()
  const input = join(dir, 'in.docx')
  const ops = join(dir, 'ops.json')
  writeFileSync(input, await plainDocx())
  writeFileSync(
    ops,
    JSON.stringify([{ op: 'wrap_content_control', startBlockIndex: 1, endBlockIndex: 99 }]),
  )

  const result = await run(['docs', 'apply', input, '--ops', ops, '--json'])
  expect(result.code).not.toBe(0)
})

it('rejects a wrap that cuts through an existing control', async () => {
  const dir = tempDir()
  const input = join(dir, 'in.docx')
  const ops = join(dir, 'ops.json')
  writeFileSync(input, await buildDocx({ bodyXml: BODY }))
  writeFileSync(
    ops,
    JSON.stringify([{ op: 'wrap_content_control', startBlockIndex: 0, endBlockIndex: 1 }]),
  )

  const result = await run(['docs', 'apply', input, '--ops', ops, '--json'])
  expect(result.code).not.toBe(0)
  expect(result.stdout + result.stderr).toMatch(/cut through/)
})

it('rejects an unwrap of a tag no control carries', async () => {
  const dir = tempDir()
  const input = join(dir, 'in.docx')
  const ops = join(dir, 'ops.json')
  writeFileSync(input, await buildDocx({ bodyXml: BODY }))
  writeFileSync(ops, JSON.stringify([{ op: 'unwrap_content_control', tag: 'example:missing' }]))

  const result = await run(['docs', 'apply', input, '--ops', ops, '--json'])
  expect(result.code).not.toBe(0)
  expect(result.stdout + result.stderr).toMatch(/no content control is tagged/)
})

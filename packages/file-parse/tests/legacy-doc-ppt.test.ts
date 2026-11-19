import { readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { docToText, parseFileToText, pptToText } from '../src/index'
import { writeFixture } from './helpers/fixtures'

function legacyFixture(name: string): string {
  return fileURLToPath(new URL(`fixtures/${name}`, import.meta.url))
}

describe('parseFileToText: legacy .doc', () => {
  it('extracts body text from a Word 97-2003 document', async () => {
    const result = await parseFileToText(legacyFixture('legacy-sample.doc'))
    expect(result.ok).toBe(true)
    expect(result.kind).toBe('text')
    expect(result.text).toContain('Legacy Report')
    expect(result.text).toContain('Legacy DOC body text')
    expect(result.text).toContain('Second paragraph from Word 97-2003.')
  })

  it('exposes the same text through docToText', async () => {
    const text = await docToText(await readFile(legacyFixture('legacy-sample.doc')))
    expect(text).toContain('Legacy Report')
    expect(text).toContain('Legacy DOC body text')
  })

  it('reads an uppercase .DOC extension through the same path', async () => {
    const path = writeFixture('upper.DOC', await readFile(legacyFixture('legacy-sample.doc')))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.text).toContain('Legacy Report')
  })

  it('rejects a corrupt .doc with an error instead of throwing', async () => {
    const path = writeFixture('corrupt.doc', Buffer.from('not a legacy document'))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(false)
    expect(result.kind).toBe('text')
    expect(result.error).toBeTruthy()
  })

  it('rejects an empty .doc with an error instead of throwing', async () => {
    const path = writeFixture('empty.doc', Buffer.alloc(0))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(false)
    expect(result.error).toBeTruthy()
  })
})

describe('parseFileToText: legacy .ppt', () => {
  it('extracts one text section per slide from a PowerPoint 97-2003 presentation', async () => {
    const result = await parseFileToText(legacyFixture('legacy-sample.ppt'))
    expect(result.ok).toBe(true)
    expect(result.kind).toBe('text')
    expect(result.text).toContain('## Slide 1')
    expect(result.text).toContain('Legacy PPT title')
    expect(result.text).toContain('First slide body')
    expect(result.text).toContain('## Slide 2')
    expect(result.text).toContain('Second legacy slide')
  })

  it('exposes the same text through pptToText', async () => {
    const text = await pptToText(await readFile(legacyFixture('legacy-sample.ppt')))
    expect(text).toContain('## Slide 1')
    expect(text).toContain('Legacy PPT title')
  })

  it('rejects a corrupt .ppt with an error instead of throwing', async () => {
    const path = writeFixture('corrupt.ppt', Buffer.from('not a legacy document'))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(false)
    expect(result.kind).toBe('text')
    expect(result.error).toBeTruthy()
  })

  it('rejects an empty .ppt with an error instead of throwing', async () => {
    const path = writeFixture('empty.ppt', Buffer.alloc(0))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(false)
    expect(result.error).toBeTruthy()
  })
})

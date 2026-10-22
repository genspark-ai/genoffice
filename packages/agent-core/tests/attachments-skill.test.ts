import { describe, expect, it, vi } from 'vitest'
import { createAttachmentsSkill, type AttachmentInfo } from '../src/attachments-skill'
import { kbCiteHref } from '../src/knowledge-base-skill'

const NOTES: AttachmentInfo = {
  path: '/tmp/my notes.md',
  name: 'my notes.md',
  ext: 'md',
  sizeBytes: 2048,
}
const PHOTO: AttachmentInfo = {
  path: '/tmp/photo.png',
  name: 'photo.png',
  ext: 'png',
  sizeBytes: 1024,
}

describe('attachments skill', () => {
  it('lists names only by default and adds paths plus cite links on request', () => {
    const plain = createAttachmentsSkill({ list: () => [NOTES], read: vi.fn() })
    expect(plain.buildContext?.()).toContain('0 | my notes.md | .md | 2KB')
    expect(plain.buildContext?.()).not.toContain('/tmp/my notes.md')
    expect(plain.systemPrompt).not.toContain('absolute paths')

    const withPaths = createAttachmentsSkill({
      list: () => [NOTES],
      read: vi.fn(),
      pathsInContext: true,
    })
    expect(withPaths.buildContext?.()).toContain('0 = /tmp/my notes.md')
    expect(withPaths.buildContext?.()).toContain(`cite: ${kbCiteHref(NOTES.path)}`)
    expect(withPaths.systemPrompt).toContain('absolute paths')
    expect(createAttachmentsSkill({ list: () => [], read: vi.fn() }).buildContext?.()).toBe('')
  })

  it('pages through the text and reports where to continue', async () => {
    const read = vi.fn(async () => ({ text: 'x'.repeat(50), totalChars: 50_000, offset: 100 }))
    const skill = createAttachmentsSkill({ list: () => [NOTES], read })
    const result = await skill.executeTool({
      id: 't1',
      name: 'read_attachment',
      input: { index: 0, offset: 100 },
    })
    expect(read).toHaveBeenCalledWith('/tmp/my notes.md', 100, 24_000)
    expect(result.isError).toBeFalsy()
    expect(result.output).toContain('this slice 100-150 (not finished, continue with offset=150)')
  })

  it('answers images locally and turns a failed read into a tool error', async () => {
    const read = vi.fn(async () => {
      throw new Error('File exceeds the size limit')
    })
    const skill = createAttachmentsSkill({ list: () => [PHOTO, NOTES], read })
    const image = await skill.executeTool({ id: 'a', name: 'read_attachment', input: { index: 0 } })
    expect(image.output).toContain('image attachment')
    expect(read).not.toHaveBeenCalled()
    const failed = await skill.executeTool({
      id: 'b',
      name: 'read_attachment',
      input: { index: 1 },
    })
    expect(failed.isError).toBe(true)
    expect(failed.output).toBe('File exceeds the size limit')
    const missing = await skill.executeTool({
      id: 'c',
      name: 'read_attachment',
      input: { index: 7 },
    })
    expect(missing.isError).toBe(true)
  })
})

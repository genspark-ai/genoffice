import { afterEach, describe, expect, it, vi } from 'vitest'

import { createImageSkill } from '../src/renderer/ai/image-skill'

function stubDesktopApi(api: Record<string, unknown>): void {
  vi.stubGlobal('window', { desktopApi: api })
}

afterEach(() => {
  vi.unstubAllGlobals()
})

function call(name: string, input: Record<string, unknown>) {
  return { id: 'call-1', name, input }
}

describe('image skill: image_search', () => {
  it('rejects an empty query', async () => {
    stubDesktopApi({})
    const result = await createImageSkill().executeTool(call('image_search', {}))
    expect(result.isError).toBe(true)
  })

  it('surfaces backend failures as errors, not empty galleries', async () => {
    stubDesktopApi({
      imageSearch: vi.fn().mockResolvedValue({ images: [], method: 'error', error: 'quota' }),
    })
    const result = await createImageSkill().executeTool(call('image_search', { query: 'cat' }))
    expect(result.isError).toBe(true)
    expect(result.output).toContain('quota')
    expect(result.output).toContain('not an empty result')
  })

  it('lists numbered direct URLs with pixel sizes', async () => {
    const imageSearch = vi.fn().mockResolvedValue({
      method: 'serper',
      images: [
        {
          title: 'Golden retriever',
          imageUrl: 'https://img.example.com/dog.jpg',
          sourceUrl: 'https://example.com',
          source: 'example',
          width: 800,
          height: 600,
        },
      ],
    })
    stubDesktopApi({ imageSearch })
    const result = await createImageSkill().executeTool(
      call('image_search', { query: 'dog', maxResults: 3 }),
    )
    expect(imageSearch).toHaveBeenCalledWith('dog', 3)
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('1. Golden retriever [800x600]')
    expect(result.output).toContain('https://img.example.com/dog.jpg')
  })
})

describe('image skill: generate_image', () => {
  it('rejects an empty prompt', async () => {
    stubDesktopApi({})
    const result = await createImageSkill().executeTool(call('generate_image', {}))
    expect(result.isError).toBe(true)
  })

  it('propagates generation errors (e.g. not logged in)', async () => {
    stubDesktopApi({
      generateImage: vi.fn().mockResolvedValue({ error: 'Genspark account is not logged in' }),
    })
    const result = await createImageSkill().executeTool(
      call('generate_image', { prompt: 'a chart mascot' }),
    )
    expect(result.isError).toBe(true)
    expect(result.output).toContain('not logged in')
  })

  it('returns the generated URL with insertion guidance', async () => {
    const generateImage = vi.fn().mockResolvedValue({ url: 'https://cdn.example.com/gen/1.png' })
    stubDesktopApi({ generateImage })
    const result = await createImageSkill().executeTool(
      call('generate_image', { prompt: 'minimal logo', aspectRatio: '1:1' }),
    )
    expect(generateImage).toHaveBeenCalledWith({ prompt: 'minimal logo', aspectRatio: '1:1' })
    expect(result.isError).toBeUndefined()
    expect(result.output).toContain('https://cdn.example.com/gen/1.png')
    expect(result.output).toContain('add_image')
  })
})

describe('image skill: unknown tool', () => {
  it('fails closed', async () => {
    stubDesktopApi({})
    const result = await createImageSkill().executeTool(call('delete_everything', {}))
    expect(result.isError).toBe(true)
  })
})

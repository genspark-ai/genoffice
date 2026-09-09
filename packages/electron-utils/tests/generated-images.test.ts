import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import { describe, expect, it } from 'vitest'
import {
  GENERATED_IMAGE_DIR,
  readGeneratedImage,
  storeGeneratedImage,
} from '../src/generated-images'
import { fetchRemoteImage } from '../src/remote-image'

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3])

describe('generated image store', () => {
  it('round-trips bytes through a file URL inside the store directory', async () => {
    const url = storeGeneratedImage(PNG, 'image/png')
    expect(url.startsWith('file://')).toBe(true)
    expect(url).toMatch(/\.png$/)
    const back = readGeneratedImage(url)
    expect(back?.mime).toBe('image/png')
    expect(Array.from(back!.bytes)).toEqual(Array.from(PNG))
    const resp = await fetchRemoteImage(url)
    expect(resp?.ok).toBe(true)
    expect(resp?.headers.get('content-type')).toBe('image/png')
    expect(new Uint8Array(await resp!.arrayBuffer())).toEqual(PNG)
  })

  it('refuses file URLs outside the store, foreign names inside it, and traversal', async () => {
    expect(readGeneratedImage('file:///etc/passwd')).toBeNull()
    const stray = join(GENERATED_IMAGE_DIR, 'notes.txt')
    writeFileSync(stray, 'x')
    expect(readGeneratedImage(pathToFileURL(stray).toString())).toBeNull()
    expect(
      readGeneratedImage(
        pathToFileURL(join(GENERATED_IMAGE_DIR, '..', 'genoffice-ai-images', 'x.png')).toString(),
      ),
    ).toBeNull()
    expect(readGeneratedImage('https://example.com/a.png')).toBeNull()
    expect(await fetchRemoteImage('file:///etc/hosts')).toBeNull()
  })
})

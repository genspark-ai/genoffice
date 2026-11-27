import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import { describe, expect, it } from 'vitest'
import { generatedImageDisplayUrl, GENERATED_IMAGE_URL_PREFIX } from '../src/generated-image-url'
import { GENERATED_IMAGE_DIR as STORE_DIR } from '../src/generated-images'
import { resolveGeneratedImageFile } from '../src/renderer-scheme'

const UUID = '12345678-1234-1234-1234-123456789abc'
const NAME = `${UUID}.png`
const fileUrl = pathToFileURL(join(STORE_DIR, NAME)).toString()

describe('generatedImageDisplayUrl', () => {
  it('maps a store file URL to a same-origin renderer path', () => {
    expect(generatedImageDisplayUrl(fileUrl)).toBe(`${GENERATED_IMAGE_URL_PREFIX}${NAME}`)
  })

  it('leaves every other URL alone', () => {
    expect(generatedImageDisplayUrl('https://example.com/a.png')).toBe('https://example.com/a.png')
    expect(generatedImageDisplayUrl('data:image/png;base64,AA==')).toBe(
      'data:image/png;base64,AA==',
    )
    expect(generatedImageDisplayUrl(pathToFileURL('/etc/passwd').toString())).toContain('file:')
    expect(
      generatedImageDisplayUrl(pathToFileURL(join(STORE_DIR, 'notes.txt')).toString()),
    ).toContain('file:')
    expect(generatedImageDisplayUrl('not a url')).toBe('not a url')
  })
})

describe('resolveGeneratedImageFile', () => {
  it('serves only store-shaped names from the renderer scheme', () => {
    const display = generatedImageDisplayUrl(fileUrl)
    expect(resolveGeneratedImageFile(`genoffice-app://slides${display}`)).toBe(
      join(STORE_DIR, NAME),
    )
    expect(resolveGeneratedImageFile(`genoffice-app://docs${display}`)).toBe(join(STORE_DIR, NAME))
  })

  it('refuses traversal and foreign names', () => {
    const prefix = 'genoffice-app://slides/generated-images/'
    for (const name of [
      'notes.txt',
      '../outside.png',
      '%2e%2e%2foutside.png',
      `${UUID}.svg`,
      `${UUID}.png/other.png`,
    ]) {
      expect(resolveGeneratedImageFile(`${prefix}${name}`)).toBeNull()
    }
    expect(resolveGeneratedImageFile(fileUrl)).toBeNull()
    expect(resolveGeneratedImageFile('genoffice-app://slides/index.html')).toBeNull()
  })
})

import { describe, expect, it } from 'vitest'
import {
  editableImageMime,
  extensionlessAssetMime,
  sniffBinaryAssetMime,
} from '../src/main/asset-mime'

const bytes = (...parts: (number[] | string)[]): Uint8Array =>
  Uint8Array.from(
    parts.flatMap((p) => (typeof p === 'string' ? [...p].map((c) => c.charCodeAt(0)) : p)),
  )

describe('sniffBinaryAssetMime', () => {
  it('recognises the common image signatures', () => {
    expect(sniffBinaryAssetMime(bytes([0xff, 0xd8, 0xff, 0xe0]))).toBe('image/jpeg')
    expect(sniffBinaryAssetMime(bytes([0x89], 'PNG', [0x0d, 0x0a, 0x1a, 0x0a]))).toBe('image/png')
    expect(sniffBinaryAssetMime(bytes('GIF89a'))).toBe('image/gif')
    expect(sniffBinaryAssetMime(bytes('RIFF', [1, 2, 3, 4], 'WEBPVP8 '))).toBe('image/webp')
    expect(sniffBinaryAssetMime(bytes([0, 0, 0, 0x1c], 'ftypavif'))).toBe('image/avif')
    expect(sniffBinaryAssetMime(bytes([0, 0, 0, 0x1c], 'ftypisom'))).toBe('video/mp4')
  })

  it('recognises fonts and audio', () => {
    expect(sniffBinaryAssetMime(bytes('wOF2'))).toBe('font/woff2')
    expect(sniffBinaryAssetMime(bytes('OTTO'))).toBe('font/otf')
    expect(sniffBinaryAssetMime(bytes([0, 1, 0, 0, 0, 12]))).toBe('font/ttf')
    expect(sniffBinaryAssetMime(bytes('true', [0, 9]))).toBe('font/ttf')
    expect(sniffBinaryAssetMime(bytes('ID3'))).toBe('audio/mpeg')
    expect(sniffBinaryAssetMime(bytes('OggS'))).toBe('audio/ogg')
  })

  it('rejects text and unknown binaries', () => {
    expect(sniffBinaryAssetMime(bytes('{"token":"secret"}'))).toBeNull()
    expect(sniffBinaryAssetMime(bytes('<!doctype html>'))).toBeNull()
    expect(sniffBinaryAssetMime(bytes('RIFF', [0, 0, 0, 0], 'AVI '))).toBeNull()
    expect(sniffBinaryAssetMime(bytes('true\n'))).toBeNull()
    expect(sniffBinaryAssetMime(bytes('BMW sales 2024, up'))).toBeNull()
    expect(sniffBinaryAssetMime(new Uint8Array())).toBeNull()
  })
})

const IMG = 'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8'
const CSS_ACCEPT = 'text/css,*/*;q=0.1'
const ANY = '*/*'

describe('extensionlessAssetMime', () => {
  const jpeg = bytes([0xff, 0xd8, 0xff, 0xe0])
  const css = bytes('@font-face { font-family: Inter; }')

  it('serves sniffed binaries whatever the name, but only into a matching slot', () => {
    expect(extensionlessAssetMime('/d/x_files/tTsEQygh', jpeg, IMG)).toBe('image/jpeg')
    expect(extensionlessAssetMime('/d/x_files/photo.php', jpeg, IMG)).toBe('image/jpeg')
    expect(extensionlessAssetMime('/d/x_files/font', bytes('wOF2'), ANY)).toBe('font/woff2')
    expect(extensionlessAssetMime('/d/x_files/tTsEQygh', jpeg, ANY)).toBeNull()
    expect(extensionlessAssetMime('/d/x_files/tTsEQygh', jpeg, CSS_ACCEPT)).toBeNull()
    expect(extensionlessAssetMime('/d/x_files/tTsEQygh', jpeg, null)).toBeNull()
    expect(extensionlessAssetMime('/d/x_files/font', bytes('wOF2'), IMG)).toBeNull()
  })

  it('types extensionless text by the requesting slot only', () => {
    expect(extensionlessAssetMime('/d/x_files/css2', css, CSS_ACCEPT)).toBe(
      'text/css; charset=utf-8',
    )
    expect(extensionlessAssetMime('/d/x_files/v31e', bytes('var a=1'), ANY)).toBe(
      'text/javascript; charset=utf-8',
    )
    expect(extensionlessAssetMime('/d/x_files/css2', css, IMG)).toBeNull()
  })

  it('never serves text neighbours that carry an extension or start with a dot', () => {
    const json = bytes('{"token":"secret"}')
    expect(extensionlessAssetMime('/d/secret.json', json, CSS_ACCEPT)).toBeNull()
    expect(extensionlessAssetMime('/d/notes.txt', json, ANY)).toBeNull()
    expect(extensionlessAssetMime('/d/.env', bytes('KEY=1'), ANY)).toBeNull()
    expect(extensionlessAssetMime('/d/flags', bytes('true\n'), IMG)).toBeNull()
  })

  it('serves an extensionless svg only into an image slot', () => {
    const svg = bytes('<?xml version="1.0"?><svg xmlns="http://www.w3.org/2000/svg"/>')
    expect(extensionlessAssetMime('/d/x_files/logo', svg, IMG)).toBe('image/svg+xml')
    expect(extensionlessAssetMime('/d/x_files/logo', svg, ANY)).toBe(
      'text/javascript; charset=utf-8',
    )
  })
})

describe('editableImageMime', () => {
  it('keeps png / gif, labels other pictures jpeg and refuses non-images', () => {
    expect(editableImageMime(bytes([0x89], 'PNG', [0x0d, 0x0a, 0x1a, 0x0a]))).toBe('image/png')
    expect(editableImageMime(bytes('GIF89a'))).toBe('image/gif')
    expect(editableImageMime(bytes([0xff, 0xd8, 0xff, 0xe0]))).toBe('image/jpeg')
    expect(editableImageMime(bytes('RIFF', [1, 2, 3, 4], 'WEBPVP8 '))).toBe('image/jpeg')
    expect(editableImageMime(bytes('wOF2'))).toBeNull()
    expect(editableImageMime(bytes('<svg xmlns="http://www.w3.org/2000/svg">'))).toBeNull()
  })
})

import { describe, expect, it } from 'vitest'
import { classifyDroppedFile, fileExt } from '../src/shared/media-kinds'

describe('classifyDroppedFile', () => {
  it('routes images by MIME type', () => {
    expect(classifyDroppedFile('photo.jpg', 'image/jpeg')).toBe('image')
    expect(classifyDroppedFile('logo.svg', 'image/svg+xml')).toBe('image')
  })

  it('routes the Insert dialog video/audio containers by extension', () => {
    for (const ext of ['mp4', 'm4v', 'mov', 'webm', 'avi']) {
      expect(classifyDroppedFile(`clip.${ext}`, '')).toBe('video')
    }
    for (const ext of ['mp3', 'wav', 'm4a', 'aac', 'ogg']) {
      expect(classifyDroppedFile(`track.${ext}`, '')).toBe('audio')
    }
    expect(classifyDroppedFile('CLIP.MOV', 'video/quicktime')).toBe('video')
  })

  it('lets the extension override an ambiguous MIME type', () => {
    expect(classifyDroppedFile('track.ogg', 'video/ogg')).toBe('audio')
    expect(classifyDroppedFile('clip.mp4', 'application/octet-stream')).toBe('video')
  })

  it('does not trust a video/audio MIME for containers addMedia cannot embed', () => {
    expect(classifyDroppedFile('clip.mkv', 'video/x-matroska')).toBe('other')
    expect(classifyDroppedFile('voice.flac', 'audio/flac')).toBe('other')
  })

  it('leaves everything else to the shell drop-open bridge', () => {
    expect(classifyDroppedFile('deck.pptx', '')).toBe('other')
    expect(classifyDroppedFile('notes.txt', 'text/plain')).toBe('other')
    expect(classifyDroppedFile('README', '')).toBe('other')
  })
})

describe('fileExt', () => {
  it('lowercases the last extension and returns empty without one', () => {
    expect(fileExt('a.b.MP4')).toBe('mp4')
    expect(fileExt('noext')).toBe('')
  })
})

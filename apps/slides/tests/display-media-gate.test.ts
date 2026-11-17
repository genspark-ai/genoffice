import { describe, expect, it } from 'vitest'
import { isSlidesScreenRecordingSource } from '../src/main/display-media-gate'

describe('isSlidesScreenRecordingSource', () => {
  it('allows the slides renderer scheme URL', () => {
    expect(isSlidesScreenRecordingSource('genoffice-app://slides/index.html', undefined)).toBe(true)
    expect(
      isSlidesScreenRecordingSource('genoffice-app://slides/index.html?mode=tab', undefined),
    ).toBe(true)
  })

  it('allows the dev-server origin when one is configured', () => {
    expect(isSlidesScreenRecordingSource('http://localhost:5175/', 'http://localhost:5175')).toBe(
      true,
    )
    expect(
      isSlidesScreenRecordingSource(
        'http://localhost:5175/index.html?x=1',
        'http://localhost:5175/index.html',
      ),
    ).toBe(true)
  })

  it('refuses other apps and hostile frames', () => {
    expect(isSlidesScreenRecordingSource('genoffice-app://html/index.html', undefined)).toBe(false)
    expect(isSlidesScreenRecordingSource('genoffice-app://docs/index.html', undefined)).toBe(false)
    expect(
      isSlidesScreenRecordingSource('html-preview://view-1/', undefined),
      'the preview iframe origin, one click from a hostile document',
    ).toBe(false)
    expect(isSlidesScreenRecordingSource('http://localhost:5177/', 'http://localhost:5175')).toBe(
      false,
    )
    expect(isSlidesScreenRecordingSource('file:///etc/passwd', undefined)).toBe(false)
    expect(isSlidesScreenRecordingSource('genoffice-app://slides.evil.com/', undefined)).toBe(false)
  })

  it('refuses garbage and absence', () => {
    expect(isSlidesScreenRecordingSource(undefined, undefined)).toBe(false)
    expect(isSlidesScreenRecordingSource('not a url', undefined)).toBe(false)
    expect(isSlidesScreenRecordingSource('', 'http://localhost:5175')).toBe(false)
  })
})

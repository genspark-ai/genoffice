import { describe, expect, it } from 'vitest'
import { audioFrame, videoFrame, videoSize } from '../src/main/video-size'

// ── EBML builders ────────────────────────────────────────────────────────

function idBytes(id: number): number[] {
  const out: number[] = []
  for (let v = id; v > 0; v = Math.floor(v / 256)) out.unshift(v & 0xff)
  return out
}

function ebml(id: number, payload: Uint8Array, unknownSize = false): Uint8Array {
  const size = unknownSize
    ? [0x01, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff, 0xff]
    : [0x80 | payload.length]
  return Uint8Array.from([...idBytes(id), ...size, ...payload])
}

function uint(id: number, value: number, len = 2): Uint8Array {
  const b: number[] = []
  for (let i = len - 1; i >= 0; i--) b.push((value >> (8 * i)) & 0xff)
  return ebml(id, Uint8Array.from(b))
}

function concat(...parts: Uint8Array[]): Uint8Array {
  return Uint8Array.from(parts.flatMap((p) => [...p]))
}

function webm(video: Uint8Array, opts: { liveSegment?: boolean } = {}): Uint8Array {
  const header = ebml(0x1a45dfa3, ebml(0x4282, Uint8Array.from([0x77, 0x65, 0x62, 0x6d])))
  const audioTrack = ebml(0xae, concat(uint(0xd7, 1, 1), uint(0x83, 2, 1)))
  const videoTrack = ebml(0xae, concat(uint(0xd7, 2, 1), uint(0x83, 1, 1), video))
  const tracks = ebml(0x1654ae6b, concat(audioTrack, videoTrack))
  const cluster = ebml(0x1f43b675, uint(0xe7, 0, 1), opts.liveSegment)
  const segment = ebml(0x18538067, concat(tracks, cluster), opts.liveSegment)
  return concat(header, segment)
}

// ── RIFF builder ─────────────────────────────────────────────────────────

function le32(n: number): number[] {
  return [n & 0xff, (n >> 8) & 0xff, (n >> 16) & 0xff, (n >>> 24) & 0xff]
}

function ascii(s: string): number[] {
  return [...s].map((c) => c.charCodeAt(0))
}

function avi(width: number, height: number): Uint8Array {
  const avih = [
    ...ascii('avih'),
    ...le32(56),
    ...new Array(32).fill(0),
    ...le32(width),
    ...le32(height),
    ...new Array(16).fill(0),
  ]
  const hdrl = [...ascii('LIST'), ...le32(4 + avih.length), ...ascii('hdrl'), ...avih]
  const junk = [...ascii('JUNK'), ...le32(3), 0, 0, 0, 0]
  const body = [...ascii('AVI '), ...junk, ...hdrl]
  return Uint8Array.from([...ascii('RIFF'), ...le32(body.length), ...body])
}

// ── Tests ────────────────────────────────────────────────────────────────

describe('videoSize', () => {
  it('reads PixelWidth/PixelHeight from a webm Video element', () => {
    const video = ebml(0xe0, concat(uint(0xb0, 1280), uint(0xba, 720)))
    expect(videoSize(webm(video), 'webm')).toEqual({ width: 1280, height: 720 })
  })

  it('prefers DisplayWidth/DisplayHeight when both are present', () => {
    const video = ebml(
      0xe0,
      concat(uint(0xb0, 720), uint(0xba, 480), uint(0x54b0, 853), uint(0x54ba, 480)),
    )
    expect(videoSize(webm(video), 'mkv')).toEqual({ width: 853, height: 480 })
  })

  it('handles MediaRecorder-style unknown-size Segment and Cluster', () => {
    const video = ebml(0xe0, concat(uint(0xb0, 1440), uint(0xba, 900)))
    expect(videoSize(webm(video, { liveSegment: true }), 'webm')).toEqual({
      width: 1440,
      height: 900,
    })
  })

  it('reads avih dimensions from an avi, skipping leading chunks', () => {
    expect(videoSize(avi(400, 300), 'avi')).toEqual({ width: 400, height: 300 })
  })

  it('returns null for unknown extensions and garbage', () => {
    expect(videoSize(avi(400, 300), 'mp3')).toBeNull()
    expect(videoSize(Uint8Array.from(ascii('nope')), 'webm')).toBeNull()
    expect(videoSize(Uint8Array.from(ascii('RIFFxxxxWAVE')), 'avi')).toBeNull()
    expect(videoSize(new Uint8Array(0), 'mp4')).toBeNull()
  })
})

describe('videoFrame', () => {
  // 13.33in x 7.5in slide = 960 x 540 pt; frames measured against PowerPoint for Mac
  const slide = { width: 12192000, height: 6858000 }
  const pt = (n: number) => n * 12700

  it('places a frame that fits at its pixel size, 1 px = 1 pt, centered', () => {
    expect(videoFrame(slide, { width: 640, height: 480 })).toEqual({
      x: pt(160),
      y: pt(30),
      cx: pt(640),
      cy: pt(480),
    })
    expect(videoFrame(slide, { width: 320, height: 180 })).toEqual({
      x: pt(320),
      y: pt(180),
      cx: pt(320),
      cy: pt(180),
    })
  })

  it('shrinks an overflowing frame to fit the slide, keeping the aspect ratio', () => {
    expect(videoFrame(slide, { width: 1920, height: 1080 })).toEqual({
      x: 0,
      y: 0,
      cx: slide.width,
      cy: slide.height,
    })
    expect(videoFrame(slide, { width: 2400, height: 600 })).toEqual({
      x: 0,
      y: pt(150),
      cx: pt(960),
      cy: pt(240),
    })
    expect(videoFrame(slide, { width: 1000, height: 400 })).toEqual({
      x: 0,
      y: pt(78),
      cx: pt(960),
      cy: pt(384),
    })
    const portrait = videoFrame(slide, { width: 1080, height: 1920 })
    expect(portrait.cy).toBe(slide.height)
    expect(portrait.cx).toBe(Math.round(pt(303.75)))
    expect(portrait.y).toBe(0)
  })

  it('centers on a drop point instead of the slide when one is given', () => {
    expect(videoFrame(slide, { width: 640, height: 480 }, { x: pt(400), y: pt(300) })).toEqual({
      x: pt(80),
      y: pt(60),
      cx: pt(640),
      cy: pt(480),
    })
    // The drop point moves the frame, never resizes it: an overflowing video still fills the slide
    const oversize = videoFrame(slide, { width: 1920, height: 1080 }, { x: pt(100), y: pt(100) })
    expect(oversize).toMatchObject({ cx: slide.width, cy: slide.height })
    expect(oversize.x).toBe(pt(100) - slide.width / 2)
    expect(videoFrame(slide, null, { x: pt(480), y: pt(270) })).toEqual(videoFrame(slide, null))
  })

  it('falls back to a 16:9 frame fitted to the slide when the size is unknown', () => {
    const fallback = { x: 0, y: 0, cx: slide.width, cy: slide.height }
    expect(videoFrame(slide, null)).toEqual(fallback)
    expect(videoFrame(slide, { width: 0, height: 0 })).toEqual(fallback)
    expect(videoFrame({ width: 9144000, height: 6858000 }, null)).toEqual({
      x: 0,
      y: 857250,
      cx: 9144000,
      cy: 5143500,
    })
  })
})

describe('audioFrame', () => {
  // PowerPoint for Mac, 960 x 540 pt slide: speaker icon at left=448 top=238, 64 x 64 pt
  it('places a 64 pt square centered on a 16:9 slide', () => {
    expect(audioFrame({ width: 12192000, height: 6858000 })).toEqual({
      x: 5689600,
      y: 3022600,
      cx: 812800,
      cy: 812800,
    })
  })

  it('centers the icon on a drop point', () => {
    expect(audioFrame({ width: 12192000, height: 6858000 }, { x: 1270000, y: 635000 })).toEqual({
      x: 863600,
      y: 228600,
      cx: 812800,
      cy: 812800,
    })
  })

  it('keeps the same 64 pt square on a 4:3 slide', () => {
    expect(audioFrame({ width: 9144000, height: 6858000 })).toEqual({
      x: 4165600,
      y: 3022600,
      cx: 812800,
      cy: 812800,
    })
  })
})

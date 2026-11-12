/**
 * MP4 writer on WebCodecs: H.264 (platform encoder) + AAC, muxed by mp4-muxer and streamed to
 * the file the main process holds open. Nothing of the finished file is kept in memory.
 */
import { Muxer, StreamTarget } from 'mp4-muxer'

export const AUDIO_SAMPLE_RATE = 48000
const KEYFRAME_EVERY = 60
const MAX_QUEUE = 8

export interface Mp4WriterOptions {
  path: string
  width: number
  height: number
  fps: number
  /** Add a mono AAC track */
  audio: boolean
}

export class EncoderUnsupportedError extends Error {}

/** Candidates from most to least capable; the first the platform accepts wins. */
const H264_CODECS = ['avc1.64002A', 'avc1.4D402A', 'avc1.42E02A']

async function pickVideoConfig(
  width: number,
  height: number,
  fps: number,
): Promise<VideoEncoderConfig> {
  const bitrate = height >= 1080 ? 8_000_000 : 5_000_000
  for (const codec of H264_CODECS) {
    const config: VideoEncoderConfig = {
      codec,
      width,
      height,
      bitrate,
      framerate: fps,
      latencyMode: 'quality',
      avc: { format: 'avc' },
    }
    const r = await VideoEncoder.isConfigSupported(config).catch(() => null)
    if (r?.supported) return config
  }
  throw new EncoderUnsupportedError('no H.264 encoder')
}

export class Mp4Writer {
  private muxer!: Muxer<StreamTarget>
  private video!: VideoEncoder
  private audio: AudioEncoder | null = null
  private fileId = 0
  private writes: Promise<void> = Promise.resolve()
  private failure: unknown = null
  private frames = 0
  private audioSamples = 0
  private readonly frameUs: number

  private constructor(private readonly opts: Mp4WriterOptions) {
    this.frameUs = Math.round(1_000_000 / opts.fps)
  }

  static async create(opts: Mp4WriterOptions): Promise<Mp4Writer> {
    const w = new Mp4Writer(opts)
    const videoConfig = await pickVideoConfig(opts.width, opts.height, opts.fps)
    let audioConfig: AudioEncoderConfig | null = null
    if (opts.audio) {
      audioConfig = {
        codec: 'mp4a.40.2',
        sampleRate: AUDIO_SAMPLE_RATE,
        numberOfChannels: 1,
        bitrate: 96_000,
      }
      const r = await AudioEncoder.isConfigSupported(audioConfig).catch(() => null)
      if (!r?.supported) audioConfig = null
    }
    w.fileId = await window.slidesApi.videoFileOpen(opts.path)
    try {
      w.openPipeline(videoConfig, audioConfig)
    } catch (e) {
      await w.abort()
      throw e
    }
    return w
  }

  private openPipeline(videoConfig: VideoEncoderConfig, audioConfig: AudioEncoderConfig | null) {
    this.muxer = new Muxer({
      target: new StreamTarget({
        chunked: true,
        chunkSize: 1 << 20,
        onData: (data, position) => this.enqueueWrite(data, position),
      }),
      video: {
        codec: 'avc',
        width: this.opts.width,
        height: this.opts.height,
        frameRate: this.opts.fps,
      },
      audio: audioConfig
        ? { codec: 'aac', sampleRate: AUDIO_SAMPLE_RATE, numberOfChannels: 1 }
        : undefined,
      fastStart: false,
    })
    this.video = new VideoEncoder({
      output: (chunk, meta) => this.muxer.addVideoChunk(chunk, meta),
      error: (e) => (this.failure ??= e),
    })
    this.video.configure(videoConfig)
    if (audioConfig) {
      this.audio = new AudioEncoder({
        output: (chunk, meta) => this.muxer.addAudioChunk(chunk, meta),
        error: (e) => (this.failure ??= e),
      })
      this.audio.configure(audioConfig)
    }
  }

  get hasAudio(): boolean {
    return this.audio != null
  }

  private enqueueWrite(data: Uint8Array, position: number) {
    // The muxer reuses its buffer: copy before the IPC call is scheduled
    const copy = data.slice()
    this.writes = this.writes
      .then(() => window.slidesApi.videoFileWrite({ id: this.fileId, data: copy, position }))
      .catch((e) => {
        this.failure ??= e
      })
  }

  private throwIfFailed() {
    if (this.failure)
      throw this.failure instanceof Error ? this.failure : new Error(String(this.failure))
  }

  /** Append one frame; `canvas` must be the output size. */
  async addFrame(canvas: HTMLCanvasElement): Promise<void> {
    this.throwIfFailed()
    if (this.video.encodeQueueSize > MAX_QUEUE) {
      await new Promise<void>((r) =>
        this.video.addEventListener('dequeue', () => r(), { once: true }),
      )
    }
    const frame = new VideoFrame(canvas, {
      timestamp: this.frames * this.frameUs,
      duration: this.frameUs,
    })
    try {
      this.video.encode(frame, { keyFrame: this.frames % KEYFRAME_EVERY === 0 })
    } finally {
      frame.close()
    }
    this.frames++
  }

  /** Append mono PCM at AUDIO_SAMPLE_RATE; consecutive calls are laid end to end. */
  async addAudio(samples: Float32Array<ArrayBuffer>): Promise<void> {
    if (!this.audio || samples.length === 0) return
    this.throwIfFailed()
    if (this.audio.encodeQueueSize > MAX_QUEUE) {
      await new Promise<void>((r) =>
        this.audio!.addEventListener('dequeue', () => r(), { once: true }),
      )
    }
    const data = new AudioData({
      format: 'f32-planar',
      sampleRate: AUDIO_SAMPLE_RATE,
      numberOfFrames: samples.length,
      numberOfChannels: 1,
      timestamp: Math.round((this.audioSamples * 1_000_000) / AUDIO_SAMPLE_RATE),
      data: samples,
    })
    try {
      this.audio.encode(data)
    } finally {
      data.close()
    }
    this.audioSamples += samples.length
  }

  /** Seconds of audio appended so far. */
  get audioSeconds(): number {
    return this.audioSamples / AUDIO_SAMPLE_RATE
  }

  async finish(): Promise<void> {
    await this.video.flush()
    if (this.audio) await this.audio.flush()
    this.throwIfFailed()
    this.muxer.finalize()
    this.video.close()
    this.audio?.close()
    await this.writes
    this.throwIfFailed()
    await window.slidesApi.videoFileClose(this.fileId, false)
  }

  /** Stop and delete the partial file. */
  async abort(): Promise<void> {
    try {
      if (this.video && this.video.state !== 'closed') this.video.close()
      if (this.audio && this.audio.state !== 'closed') this.audio.close()
    } catch {
      // already closed by a failure
    }
    await this.writes.catch(() => undefined)
    await window.slidesApi.videoFileClose(this.fileId, true).catch(() => undefined)
  }
}

/** The presenter's microphone over a whole slide show: one continuous recording, pausable. */
export class NarrationRecorder {
  private readonly chunks: Blob[] = []
  private recordedMs = 0
  private runningSince: number | null = null

  private constructor(
    private readonly stream: MediaStream,
    private readonly recorder: MediaRecorder,
  ) {
    recorder.ondataavailable = (e) => {
      if (e.data.size > 0) this.chunks.push(e.data)
    }
  }

  static async start(): Promise<NarrationRecorder> {
    const stream = await navigator.mediaDevices.getUserMedia({
      audio: { echoCancellation: true, noiseSuppression: true },
    })
    const recording = new NarrationRecorder(stream, new MediaRecorder(stream))
    // A timeslice keeps what was said if the show ends abruptly
    recording.recorder.start(1000)
    recording.runningSince = performance.now()
    return recording
  }

  /** Recorded time so far, pauses excluded. */
  elapsedMs(): number {
    return (
      this.recordedMs + (this.runningSince === null ? 0 : performance.now() - this.runningSince)
    )
  }

  get paused(): boolean {
    return this.runningSince === null
  }

  pause(): void {
    if (this.runningSince === null || this.recorder.state !== 'recording') return
    this.recordedMs = this.elapsedMs()
    this.runningSince = null
    this.recorder.pause()
  }

  resume(): void {
    if (this.runningSince !== null || this.recorder.state !== 'paused') return
    this.runningSince = performance.now()
    this.recorder.resume()
  }

  stop(): Promise<{ audio: Blob; endMs: number }> {
    const endMs = this.elapsedMs()
    this.runningSince = null
    this.recordedMs = endMs
    return new Promise((resolve) => {
      const done = () => {
        this.dispose()
        resolve({ audio: new Blob(this.chunks, { type: this.recorder.mimeType }), endMs })
      }
      if (this.recorder.state === 'inactive') return done()
      this.recorder.addEventListener('stop', done, { once: true })
      this.recorder.stop()
    })
  }

  dispose(): void {
    for (const track of this.stream.getTracks()) track.stop()
  }
}

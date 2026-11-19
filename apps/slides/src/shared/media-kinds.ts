/** Media containers the Insert dialog offers; a drop of the same files inserts them the same way. */
export const VIDEO_EXTS = ['mp4', 'm4v', 'mov', 'webm', 'avi'] as const
export const AUDIO_EXTS = ['mp3', 'wav', 'm4a', 'aac', 'ogg'] as const

export type DroppedFileKind = 'image' | 'video' | 'audio' | 'other'

export function fileExt(name: string): string {
  const dot = name.lastIndexOf('.')
  return dot < 0 ? '' : name.slice(dot + 1).toLowerCase()
}

/**
 * Classify a dropped file the way the canvas inserts it. Media goes by extension
 * alone: the lists are exactly the containers addMedia can embed, and Chromium's
 * MIME for them is unreliable (ogg/mov vary, m4v is empty). Anything else must
 * stay 'other' so the drop reaches the shell's drop-open bridge.
 */
export function classifyDroppedFile(name: string, mime: string): DroppedFileKind {
  const ext = fileExt(name)
  if ((VIDEO_EXTS as readonly string[]).includes(ext)) return 'video'
  if ((AUDIO_EXTS as readonly string[]).includes(ext)) return 'audio'
  if (mime.startsWith('image/')) return 'image'
  return 'other'
}

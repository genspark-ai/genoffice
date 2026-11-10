import type { ImageHostConfig } from '../shared/ipc'

// The paste path needs the host config synchronously-ish but must not ask the
// main process on every paste; the dialog refreshes the cache on save so the
// next paste uses the new configuration at once. A failed read keeps the cache
// unset and retries on the next paste (the dialog is the only writer).
let cached: { config: ImageHostConfig | null } | null = null

export async function getImageHostConfig(): Promise<ImageHostConfig | null> {
  if (cached) return cached.config
  try {
    const config = await window.markdownApi.getImageHost()
    cached = { config }
    return config
  } catch {
    return null
  }
}

export function setImageHostConfigCache(config: ImageHostConfig | null): void {
  cached = { config }
}

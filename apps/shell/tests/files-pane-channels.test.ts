import { describe, expect, it } from 'vitest'
import { HOME_CHANNELS } from '../src/shared/home-api'
import { FILES_PANE_CHANNELS } from '@genoffice/electron-utils/files-pane-bridge'

/** the editors' Files pane bridge hardcodes the shell's home channels; keep them identical */
describe('files pane bridge channels', () => {
  it('match HOME_CHANNELS', () => {
    for (const [key, channel] of Object.entries(FILES_PANE_CHANNELS)) {
      expect(HOME_CHANNELS[key as keyof typeof HOME_CHANNELS]).toBe(channel)
    }
  })
})

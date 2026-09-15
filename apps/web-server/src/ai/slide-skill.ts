/**
 * Slide-AI skill handlers — outline/content/full generation, style apply /
 * consistency, notes generation, translate.
 *
 * Web channels intentionally throw `WebUnsupportedError`: editing happens
 * in the renderer-side skill (`apps/slides/src/renderer/ai/slides-skill.ts`),
 * which talks to the same model via the slides app's AI panel. The
 * web-server is the IPC + storage layer; it does not fabricate AI output.
 */
import { registerHandle } from '../common/index'
import { WebUnsupportedError } from './errors'

const WEB_SLIDE_CHANNELS = [
  'ai:slides-generate-outline',
  'ai:slides-generate-content',
  'ai:slides-generate-full',
  'ai:slides-style-apply',
  'ai:slides-style-consistency',
  'ai:slides-generate-notes',
  'ai:slides-translate',
] as const

export function registerSlideAiSkillHandlers(): void {
  for (const channel of WEB_SLIDE_CHANNELS) {
    registerHandle(channel, () => {
      throw new WebUnsupportedError(channel, 'renderer-side skill')
    })
  }
}

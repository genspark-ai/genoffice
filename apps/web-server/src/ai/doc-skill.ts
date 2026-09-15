/**
 * Doc-AI skill handlers — write continue/expand/shrink/rewrite, tone adjust,
 * format suggest/apply.
 *
 * Web channels intentionally throw `WebUnsupportedError`: editing happens
 * in the renderer-side skill (`apps/docs/src/renderer/ai/docs-skill.ts`),
 * which talks to the same model via the docs app's AI panel. The
 * web-server is the IPC + storage layer; it does not fabricate AI output.
 */
import { registerHandle } from '../common/index'
import { WebUnsupportedError } from './errors'

const WEB_DOC_CHANNELS = [
  'ai:doc-write-continue',
  'ai:doc-write-expand',
  'ai:doc-write-shrink',
  'ai:doc-write-rewrite',
  'ai:doc-tone-adjust',
  'ai:doc-format-suggest',
  'ai:doc-format-apply',
] as const

export function registerDocAiSkillHandlers(): void {
  for (const channel of WEB_DOC_CHANNELS) {
    registerHandle(channel, () => {
      throw new WebUnsupportedError(channel, 'renderer-side skill')
    })
  }
}

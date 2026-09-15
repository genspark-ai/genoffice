/**
 * Sheet-AI skill handlers — formula suggest/explain, data analyze,
 * trend predict, chart suggest/create, data clean/fill.
 *
 * Web channels intentionally throw `WebUnsupportedError`: editing happens
 * in the renderer-side skill (`apps/sheets/src/renderer/ai/workbook-skill.ts`),
 * which talks to the same model via the sheets app's AI panel. The
 * web-server is the IPC + storage layer; it does not fabricate AI output.
 */
import { registerHandle } from '../common/index'
import { WebUnsupportedError } from './errors'

const WEB_SHEET_CHANNELS = [
  'ai:sheets-formula-suggest',
  'ai:sheets-formula-explain',
  'ai:sheets-data-analyze',
  'ai:sheets-trend-predict',
  'ai:sheets-chart-suggest',
  'ai:sheets-chart-create',
  'ai:sheets-data-clean',
  'ai:sheets-data-fill',
] as const

export function registerSheetAiSkillHandlers(): void {
  for (const channel of WEB_SHEET_CHANNELS) {
    registerHandle(channel, () => {
      throw new WebUnsupportedError(channel, 'renderer-side skill')
    })
  }
}

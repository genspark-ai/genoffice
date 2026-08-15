import { describe, expect, it } from 'vitest'
import { LANGS } from '@genoffice/i18n'
import type { IpcErrorCode } from '@genoffice/agent-core'
import {
  codexStrings,
  getCodexUiLabels,
  resolveCodexError,
  type CodexReasoningEffort,
} from '../src'

const reasoningEfforts: CodexReasoningEffort[] = [
  'none',
  'minimal',
  'low',
  'medium',
  'high',
  'xhigh',
  'max',
  'ultra',
]

const errorCodes: IpcErrorCode[] = [
  'timeout',
  'credits',
  'auth-required',
  'auth-expired',
  'auth-temporary',
  'capabilities-unavailable',
  'rate-limit',
  'request-rejected',
  'invalid-stream',
  'invalid-tool-call',
  'provider-failure',
]

describe('Codex localization', () => {
  it('covers every catalog key for every supported locale', () => {
    const keys = Object.keys(codexStrings.zh) as Array<keyof typeof codexStrings.zh>

    expect(keys.length).toBeGreaterThan(0)
    for (const lang of LANGS) {
      for (const key of keys) {
        expect(codexStrings[lang][key]).toEqual(expect.any(String))
        expect(codexStrings[lang][key].trim()).not.toBe('')
      }
    }
  })

  it('localizes every reasoning effort in every supported locale', () => {
    for (const lang of LANGS) {
      const labels = getCodexUiLabels(lang)
      for (const effort of reasoningEfforts) {
        expect(labels.reasoningLabel(effort)).not.toBe('')
      }
    }
  })

  it('resolves every IPC error code and keeps auth failures distinct', () => {
    for (const lang of LANGS) {
      for (const code of errorCodes) {
        expect(resolveCodexError(code, lang)).not.toBe('')
      }
      expect(resolveCodexError('unrecognized-code', lang)).not.toBe('')
    }

    expect(resolveCodexError('auth-required', 'en')).not.toBe(
      resolveCodexError('auth-expired', 'en'),
    )
    expect(resolveCodexError('auth-expired', 'en')).not.toBe(
      resolveCodexError('auth-temporary', 'en'),
    )
  })
})

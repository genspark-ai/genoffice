/**
 * Settings › "show the AI sidebar" is a soft hide, the way Word treats the
 * Copilot pane: an explicit Ask-AI action still opens the panel, but turning
 * the switch off closes it. The rule itself lives in one shared hook
 * (`useAiSidebarEnabled`, unit-tested in packages/ui) — otherwise the defect
 * of the previous review comes straight back: a panel left open behind an off
 * switch, with every Ask-AI entry point reduced to a no-op until a reload.
 *
 * Mounting six editors is out of reach in one suite, so what is assertable
 * here is the wiring: each editor reads the switch through that hook and hands
 * it the setter for its own panel state. An editor that dropped the call, or
 * that passed a setter for something else, would strand its panel — so every
 * call site is pinned by name, along with the derived-state workaround
 * (`aiOpenedExplicitly`) that the hook makes unnecessary.
 */
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '../../..')

/** Every editor that owns an AI dock, with the panel setter it owns. */
const EDITORS = [
  { app: 'docs', file: 'apps/docs/src/renderer/App.tsx', setter: 'setShowAi' },
  { app: 'slides', file: 'apps/slides/src/renderer/App.tsx', setter: 'setShowAi' },
  { app: 'html', file: 'apps/html/src/renderer/App.tsx', setter: 'setAiOpen' },
  { app: 'markdown', file: 'apps/markdown/src/renderer/App.tsx', setter: 'setAiOpen' },
  { app: 'pdf', file: 'apps/pdf/src/renderer/App.tsx', setter: 'setAiOpen' },
  { app: 'sheets', file: 'apps/sheets/src/renderer/ExcelShell.tsx', setter: 'setIsCopilotOpen' },
]

const source = (file: string) => readFileSync(join(repoRoot, file), 'utf8')

describe('the AI sidebar switch reaches every editor', () => {
  it.each(EDITORS)('$app closes its panel through the shared hook', ({ file, setter }) => {
    const code = source(file)
    expect(code).toContain('useAiSidebarEnabled')
    expect(code).toContain(`useAiSidebarEnabled(${setter})`)
  })

  it.each(EDITORS)(
    '$app no longer gates the panel on a derived "opened explicitly" flag',
    ({ file }) => {
      // the old derivation, `showAi && (enabled || openedExplicitly)`, is what
      // stranded docs and slides: no transition, so the flag never flipped
      const code = source(file)
      expect(code).not.toMatch(/aiOpenedExplicitly|prevShowAiRef/)
    },
  )
})

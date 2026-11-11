/** @vitest-environment jsdom */
import { spawn } from 'node:child_process'
import { describe, expect, it } from 'vitest'

import { runCopyEmfHelper } from '../src/main/copy-emf-helper'
import { copyEmfCanvasTooLarge } from '../src/renderer/copy-emf'

describe('runCopyEmfHelper', () => {
  it('a helper that dies before reading the payload does not take the process down', async () => {
    // The write fails with EPIPE once the child is gone. Without a listener on
    // the stdin stream that error is unhandled, and it takes the main process
    // down with it — the copy's text formats are already on the clipboard by
    // then, so the app must survive it.
    const result = await runCopyEmfHelper('unused-helper', 'x'.repeat(1_000_000), {
      timeoutMs: 5_000,
      // a process that exits at once, and a payload bigger than the pipe
      // buffer: the write cannot finish and fails with EPIPE
      spawnFn: ((_command: string, _args: readonly string[], options: never) =>
        spawn(process.execPath, ['-e', 'process.exit(0)'], options)) as never,
    })
    expect(result.ok).toBe(false)
  })
})

describe('copyEmfCanvasTooLarge', () => {
  it('refuses a canvas past the helper limit', () => {
    expect(copyEmfCanvasTooLarge([100_001], [10])).toBe(true)
    expect(copyEmfCanvasTooLarge([10], [100_001])).toBe(true)
    expect(copyEmfCanvasTooLarge([50_000, 50_000], [10])).toBe(false)
    expect(copyEmfCanvasTooLarge([], [])).toBe(false)
  })
})

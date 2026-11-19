import { test, expect } from '@playwright/test'
import { spawn } from 'node:child_process'
import { once } from 'node:events'
import { closeAndSaveVideo, type LaunchedApp } from './helpers'

for (const ignoreTerminate of [false, true]) {
  test(`app cleanup waits for process exit (ignores SIGTERM: ${ignoreTerminate})`, async () => {
    test.skip(process.platform === 'win32', 'Requires POSIX signal handling')
    const child = spawn(process.execPath, [
      '-e',
      `${ignoreTerminate ? "process.on('SIGTERM', () => {});" : ''}
       setInterval(() => {}, 1000);
       process.stdout.write('ready');`,
    ])
    const closed = once(child, 'close')
    try {
      await once(child.stdout!, 'data')
      // Use a real child process; only the Electron dialog/video APIs are omitted.
      const launched = {
        app: {
          evaluate: async () => {},
          process: () => child,
          close: async () => {
            child.kill('SIGTERM')
            await closed
          },
        },
        page: { video: () => null },
      } as unknown as LaunchedApp

      await closeAndSaveVideo(launched, 'cleanup')
      expect(child.exitCode !== null || child.signalCode !== null).toBe(true)
    } finally {
      if (child.exitCode === null && child.signalCode === null) child.kill('SIGKILL')
      await closed
    }
  })
}

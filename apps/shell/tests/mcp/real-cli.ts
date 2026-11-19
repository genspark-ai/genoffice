import { runCli } from '@genoffice/cli'
import type {
  CliRunner,
  CliRunOutcome,
  CliJsonOk,
  CliJsonError,
} from '../../src/main/mcp/cli-runner'

/**
 * A CliRunner backed by the real `@genoffice/cli` running in-process.
 *
 * This is the counterpart to `fakeCli`: it exercises the actual argv contract
 * (the flags the MCP tools build against the real commands) without spawning a
 * process or depending on a built `dist/genoffice.cjs`, so it works in CI under
 * plain `npm test`. It captures the CLI's `--json` stdout and maps it to the
 * same outcome shape the spawn runner returns.
 */
export function realCliRunner(cwd?: string): CliRunner {
  return {
    async run(args): Promise<CliRunOutcome> {
      let stdout = ''
      let stderr = ''
      // the spawn runner appends --json; mirror that so the contract is identical
      const code = await runCli([...args, '--json'], {
        ...(cwd ? { cwd } : {}),
        io: {
          stdout: (text) => (stdout += `${text}\n`),
          stderr: (text) => (stderr += `${text}\n`),
        },
      })
      const trimmed = stdout.trim()
      let json: CliJsonOk | CliJsonError | undefined
      try {
        const parsed = JSON.parse(trimmed.split('\n').filter(Boolean).at(-1) ?? '') as
          CliJsonOk | CliJsonError
        json = parsed && typeof parsed === 'object' && 'status' in parsed ? parsed : undefined
      } catch {
        json = undefined
      }
      return { ok: code === 0 && json?.status === 'ok', code, json, stdout, stderr }
    },
  }
}

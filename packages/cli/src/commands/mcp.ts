import { flagString } from '../args'
import type { CommandDef } from '../registry'
import { CliError, EXIT } from '../result'

export const mcpCommand: CommandDef = {
  name: 'mcp',
  summary:
    'Serve every command as a Model Context Protocol tool: on stdio for an MCP client on this machine (`claude mcp add --transport stdio genoffice -- genoffice mcp`, Cursor, Claude Desktop), or with --http as a Streamable HTTP server that clients on other machines reach by URL. Over HTTP, files travel too: PUT /files/<name> uploads one and every tool takes http(s) URLs in place of paths; outputs come back as download URLs and, when small, as embedded resources. Ops, specs and Markdown are passed inline either way; a new deck goes through deck_start, deck_page and deck_build.',
  usage: 'mcp [--http <port> [--host <addr>] [--token <secret>]]',
  quiet: true,
  options: [
    {
      name: 'http',
      value: 'port',
      description: 'serve Streamable HTTP on this port instead of stdio (0 picks a free port)',
    },
    {
      name: 'host',
      value: 'addr',
      description: 'address to bind with --http (default: 127.0.0.1; 0.0.0.0 for other machines)',
    },
    {
      name: 'token',
      value: 'secret',
      description:
        'with --http, require this bearer token on every request (default: $GENOFFICE_MCP_TOKEN, else none)',
    },
  ],
  async run(args, ctx) {
    const http = flagString(args, 'http')
    if (http === undefined) {
      // lazy: the server imports runCli, which registers this command
      const { serveStdio } = await import('../mcp/server')
      await serveStdio({ cwd: ctx.cwd, env: ctx.env, log: ctx.log })
      return { summary: 'mcp session ended' }
    }
    const port = Number(http)
    if (!Number.isInteger(port) || port < 0 || port > 65535) {
      throw new CliError(EXIT.usage, `--http needs a port number, got ${http}`, undefined, {
        reason: 'invalid_argument',
      })
    }
    const tokenFlag = flagString(args, 'token')
    // `--token ""` — e.g. `--token "$MY_TOKEN"` with the variable unset —
    // would reach authorized() as "no auth configured": the operator believes
    // the lock is on while every request passes. Fail closed instead.
    if (tokenFlag !== undefined && tokenFlag.trim() === '') {
      throw new CliError(
        EXIT.usage,
        '--token needs a non-empty value (omit the flag, or set GENOFFICE_MCP_TOKEN)',
        undefined,
        { reason: 'invalid_argument' },
      )
    }
    const { serveHttp } = await import('../mcp/http')
    await serveHttp({
      cwd: ctx.cwd,
      env: ctx.env,
      log: ctx.log,
      port,
      host: flagString(args, 'host'),
      token: tokenFlag ?? (ctx.env.GENOFFICE_MCP_TOKEN?.trim() || undefined),
    })
    return { summary: 'mcp server stopped' }
  },
}

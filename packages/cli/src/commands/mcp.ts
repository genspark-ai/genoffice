import type { CommandDef } from '../registry'

export const mcpCommand: CommandDef = {
  name: 'mcp',
  summary:
    'Serve every command as a Model Context Protocol tool on stdio, for an MCP client such as Claude Code (`claude mcp add --transport stdio genoffice -- genoffice mcp`), Cursor or Claude Desktop. Ops, specs and Markdown are passed inline; a new deck goes through deck_start, deck_page and deck_build.',
  usage: 'mcp',
  quiet: true,
  async run(_args, ctx) {
    // lazy: the server imports runCli, which registers this command
    const { serveStdio } = await import('../mcp/server')
    await serveStdio({ cwd: ctx.cwd, env: ctx.env, log: ctx.log })
    return { summary: 'mcp session ended' }
  },
}

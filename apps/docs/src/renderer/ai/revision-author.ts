/**
 * AI authorship: every edit the app makes on its own behalf carries this
 * author — the in-app assistant's tracked replacements and comment replies,
 * and the in-app MCP server's content tools, which all route through the
 * same executor this module feeds (#2654). The review pane colors the AI
 * author like any other non-own author, and the writing view's four-way
 * "who wrote what" split keys off this name.
 */
export const AI_AUTHOR = 'AI'

export function aiRevisionAuthor(): string {
  return AI_AUTHOR
}

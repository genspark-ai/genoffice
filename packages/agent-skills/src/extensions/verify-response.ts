/**
 * verify-response extension — appends GenOffice's claim/action-consistency
 * verification rules to the model system prompt at `before_agent_start`.
 *
 * Why this exists (parity with GenOffice legacy behaviour):
 *   The legacy `AgentLoop` calls `skill.verifyResponse(text, executedTools)`
 *   after each assistant message. If it returns a non-null correction, the
 *   loop re-prompts the model with that correction. On pi, the agent loop
 *   doesn't expose that hook directly; the closest seam is to inject the
 *   verification rules into the system prompt so the model self-checks
 *   before emitting a claim.
 *
 * Mechanism:
 *   pi's `BeforeAgentStartEventResult.systemPrompt` replaces the entire
 *   prompt. Multiple extensions' results chain — each sees the previous
 *   extension's output as `event.systemPrompt`. We append a `[Verify]`
 *   section to whatever prompt is in flight.
 */

import type { ExtensionAPI } from "@earendil-works/pi-coding-agent";

/**
 * The default verification rules — translated from GenOffice's
 * `packages/agent-core/src/verify-response.ts` heuristics into prose that
 * the model reads once per turn.
 *
 * Hosts can pass their own rules via `rules` if they want different
 * constraints (slides have different action vocabulary from docs).
 */
export const DEFAULT_VERIFY_RULES = `\
[Verify rules — applied every turn]
1. Only claim an action if you have just emitted the corresponding tool call this turn.
2. If you mention a result ("I have done X"), the matching tool must appear in the
   previous tool_calls; otherwise say "I will …" instead.
3. Don't claim block indices / ranges you have not actually read with read_blocks.
4. Don't promise changes outside the user's stated scope. If unsure, ask first.
5. If a tool returned an error, surface the error verbatim — never paraphrase
   a failure as success.
6. Each turn: at most one write tool per tool batch unless the user explicitly
   asked for chained writes.`;

/**
 * Marker that wraps the injected rules so it's easy to detect / strip during tests
 * and so the model can tell the section apart from the host system prompt.
 */
export const VERIFY_BLOCK_MARKER = "[genoffice:verify-rules]";

export interface VerifyResponseOptions {
  /**
   * Rules text to inject. Defaults to `DEFAULT_VERIFY_RULES`. May contain
   * template tokens like `{editor}` that the host wants to fill in.
   */
  rules?: string;
  /**
   * Marker prefix so downstream tools can recognise the injected section.
   * Defaults to `VERIFY_BLOCK_MARKER`. Pass a custom string if the host
   * already uses a similar convention.
   */
  marker?: string;
}

function formatBlock(marker: string, rules: string): string {
  return `\n\n${marker}\n${rules}\n${marker}\n`;
}

/**
 * Create the pi extension factory.
 */
export function createVerifyResponseExtension(opts: VerifyResponseOptions = {}) {
  const rules = opts.rules ?? DEFAULT_VERIFY_RULES;
  const marker = opts.marker ?? VERIFY_BLOCK_MARKER;
  const block = formatBlock(marker, rules);

  return (pi: ExtensionAPI): void => {
    pi.on("before_agent_start", async (event) => {
      // Mutate a copy — `event.systemPrompt` is read by the runner but is
      // safe to leave untouched in case other extensions want the original.
      const next = `${event.systemPrompt}${block}`;
      return { systemPrompt: next };
    });
  };
}

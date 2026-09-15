/**
 * agent-team extension — register a small registry of "review roles" and a
 * `request_review` tool the main writer can invoke to spawn an inner review
 * turn on the conversation so far.
 *
 * Why this exists (parity with GenOffice legacy behaviour):
 *   The legacy `packages/agent-core/src/agent-loop.ts` exposed a private
 *   "reviewer skill" that auto-fired after every assistant message to
 *   catch claim/action mismatches and numerical drift. On pi we expose the
 *   same idea but as opt-in tool calls so the writer (or the host's
 *   `verify-response` extension) decides when a review is needed.
 *
 * Mechanism:
 *   - The host installs the extension via `installAgentTeam(pi, opts)`.
 *   - The role registry (defaults: `reviewer`, `fact-checker`, `editor`,
 *     `summarizer`) carries a short `systemPrompt` snippet.
 *   - `request_review({ role, focus? })` calls `pi.sendUserMessage(...)`
 *     with a message that asks the model to read the conversation so far
 *     and respond in the chosen role's voice.
 *   - `sendUserMessage` keeps the result in the same session, so the writer
 *     can read the reviewer's findings and react on the next turn.
 *
 * Note: this extension does NOT spawn a true sub-agent (pi does not yet
 * expose a subagent API to extensions in 0.85.1). The reviewer is a regular
 * model turn with a constrained system prompt. This matches pi's own
 * "agents/" subagent example pattern.
 */

import { Type, type Static } from "@earendil-works/pi-ai";
import { defineTool, type ExtensionAPI } from "@earendil-works/pi-coding-agent";

/* ------------------------------------------------------------------ */
/* Role registry.                                                       */
/* ------------------------------------------------------------------ */

export interface AgentRole {
	/** Stable identifier (matches `request_review`'s `role` argument). */
	name: string;
	/** Human-readable summary shown to the model. */
	description: string;
	/**
	 * Short system prompt snippet dispatched with the reviewer turn.
	 * Keep it terse — the model already has the full conversation context.
	 */
	systemPrompt: string;
}

/** Default writer / reviewer / fact-checker roles from plan §5.2. */
export const BUILTIN_AGENT_ROLES: Readonly<Record<string, AgentRole>> = Object.freeze({
	writer: {
		name: "writer",
		description: "Default role — produces the answer to the user's request.",
		systemPrompt:
			"You are the writer. Produce the next reply that directly answers the user's most recent request.",
	},
	reviewer: {
		name: "reviewer",
		description: "Silent reviewer — claims / action consistency only.",
		systemPrompt:
			"You are the reviewer. Read the writer's last assistant turn. Reply with: " +
			"(1) any claim that is not backed by a tool call this turn, " +
			"(2) any missing action that the writer promised but did not perform, " +
			"(3) any result paraphrased in a way that hides an error. " +
			"Be specific: quote the offending sentence, then explain. " +
			"If nothing is wrong, reply exactly 'No issues.' and stop.",
	},
	fact_checker: {
		name: "fact_checker",
		description: "Numerical accuracy — checks every number against the conversation.",
		systemPrompt:
			"You are the fact-checker. Re-read the writer's last turn and the tool results it cites. " +
			"Verify every number (token counts, percentages, currency amounts, IDs, dates) against the actual tool output. " +
			"List any number that does not match exactly, quoting the writer's text and the tool's text. " +
			"If everything checks out, reply exactly 'Numbers match.' and stop.",
	},
	editor: {
		name: "editor",
		description: "Style / clarity pass on the writer's last turn.",
		systemPrompt:
			"You are the editor. Read the writer's last turn and suggest concrete edits to improve clarity and brevity. " +
			"Reply with a numbered list of 'original → revised' pairs. Do not change technical content. " +
			"If no edits are needed, reply exactly 'No edits.' and stop.",
	},
	summarizer: {
		name: "summarizer",
		description: "Compresses the conversation into a brief for the host UI.",
		systemPrompt:
			"You are the summarizer. Produce a 1-3 sentence summary of the conversation so far suitable for a session list entry. " +
			"Focus on what was done and what is left to do. Do not add commentary.",
	},
});

/* ------------------------------------------------------------------ */
/* Tool schema.                                                         */
/* ------------------------------------------------------------------ */

export const RequestReviewParams = Type.Object({
	role: Type.String({
		description:
			"Which reviewer role should run (writer / reviewer / fact_checker / editor / summarizer, or a host-registered role name).",
	}),
	focus: Type.Optional(
		Type.String({
			description: "Optional one-line instruction that overrides the role's default scope for this call.",
		}),
	),
});

export type RequestReviewArgs = Static<typeof RequestReviewParams>;

export interface RequestReviewDetails {
	role: string;
	resolvedRoleName: string;
	deliveredAs: "followUp" | "steer";
}

export interface AgentTeamOptions {
	/**
	 * Additional or replacement roles. Keys become the legal values of the
	 * `role` argument; defaults from {@link BUILTIN_AGENT_ROLES} are merged in
	 * and overridden by host entries with the same name.
	 */
	additionalRoles?: Readonly<Record<string, AgentRole>>;
	/**
	 * How the reviewer message is queued against the writer's in-flight turn.
	 * Defaults to `followUp`, which lets the writer finish first.
	 */
	deliveredAs?: "followUp" | "steer";
}

/* ------------------------------------------------------------------ */
/* Tool factory.                                                        */
/* ------------------------------------------------------------------ */

export function createRequestReviewTool(
	opts: AgentTeamOptions & {
		roles: Readonly<Record<string, AgentRole>>;
		/** Bound sender used to dispatch the review turn. Wired by `installAgentTeam`. */
		sendUserMessage: (
			content: string,
			options?: { deliverAs?: "followUp" | "steer" },
		) => Promise<void>;
	},
) {
	const { roles, sendUserMessage } = opts;
	const deliveredAs = opts.deliveredAs ?? "followUp";

	return defineTool<typeof RequestReviewParams, RequestReviewDetails>({
		name: "request_review",
		label: "Request Review",
		description:
			"Spawn an inner review turn on the conversation so far using the named role. Use after writing a draft to catch claim/action mismatches, verify numbers, tighten style, or summarize. " +
			"Each role has its own short system prompt; the model reads the full conversation context and replies in that role's voice.",
		promptSnippet:
			"request_review(role, focus?) — review the writer's last turn in a chosen voice (reviewer / fact_checker / editor / summarizer)",
		promptGuidelines: [
			"Call this after completing a draft (write tool succeeded) — never mid-write.",
			"`reviewer` is for claim/action consistency; `fact_checker` is for numbers; `editor` is for prose; `summarizer` is for one-line briefs.",
			"`focus` overrides the role's default scope for one call without mutating the registry.",
			"You can chain calls: `request_review('reviewer')` then `request_review('fact_checker')`.",
		],
		parameters: RequestReviewParams,
		async execute(
			_toolCallId,
			params: RequestReviewArgs,
			_signal,
			_onUpdate,
			ctx,
		) {
			const role = roles[params.role];
			if (!role) {
				const known = Object.keys(roles).sort().join(", ");
				throw new Error(`Unknown review role "${params.role}". Known roles: ${known}`);
			}

			const focusLine = params.focus?.trim();
			const intro = focusLine
				? `[role:${role.name}] ${focusLine}`
				: `[role:${role.name}] ${role.systemPrompt}`;
			await sendUserMessage(intro, { deliverAs: deliveredAs });

			return {
				content: [
					{
						type: "text",
						text: `Dispatched ${role.name} review${focusLine ? ` with focus: ${focusLine}` : ""}.`,
					},
				],
				details: {
					role: params.role,
					resolvedRoleName: role.name,
					deliveredAs,
				},
			};
		},
	});
}

/* ------------------------------------------------------------------ */
/* Extension installer.                                                 */
/* ------------------------------------------------------------------ */

export interface InstallAgentTeamOptions extends AgentTeamOptions {
	/** Optional human-readable name shown in extension listings. */
	extensionName?: string;
}

/**
 * Register the `request_review` tool on a pi extension API.
 * Merges host-supplied roles on top of the built-in defaults.
 */
export function installAgentTeam(pi: ExtensionAPI, opts: InstallAgentTeamOptions = {}): void {
	const merged: Record<string, AgentRole> = { ...BUILTIN_AGENT_ROLES };
	if (opts.additionalRoles) {
		for (const [name, role] of Object.entries(opts.additionalRoles)) {
			merged[name] = { ...role, name };
		}
	}
	const tool = createRequestReviewTool({
		roles: merged,
		deliveredAs: opts.deliveredAs,
		sendUserMessage: (content, options) =>
			pi.sendUserMessage(content, options ?? undefined) as unknown as Promise<void>,
	});
	pi.registerTool(tool);
}

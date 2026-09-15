/**
 * Tests for the agent-team review-roles extension (W14 deliverable).
 *
 * Verifies:
 *   - Default role registry has the five plan §5.2 roles
 *   - createRequestReviewTool dispatches via sendUserMessage with the role snippet
 *   - focus override prefixes the role line
 *   - Unknown role throws with the known-roles list
 *   - installAgentTeam wires a single tool into the ExtensionAPI
 *   - Host can override a built-in role and add a new one
 */

import { describe, expect, it, vi } from "vitest";
import type { ExtensionAPI, ExtensionContext } from "@earendil-works/pi-coding-agent";
import {
	BUILTIN_AGENT_ROLES,
	createRequestReviewTool,
	installAgentTeam,
	type AgentRole,
} from "../src/extensions/agent-team";

interface Harness {
	api: ExtensionAPI;
	captured: {
		registerTool: ReturnType<typeof vi.fn>;
		sendUserMessage: ReturnType<typeof vi.fn>;
	};
	ctx: ExtensionContext;
}

function makeHarness(): Harness {
	const captured = {
		registerTool: vi.fn(),
		sendUserMessage: vi.fn(async () => undefined),
	};
	const api = {
		registerTool: captured.registerTool,
		sendUserMessage: captured.sendUserMessage,
	} as unknown as ExtensionAPI;
	const ctx = {} as unknown as ExtensionContext;
	return { api, captured, ctx };
}

function makeTool(captured: Harness["captured"]) {
	return createRequestReviewTool({
		roles: BUILTIN_AGENT_ROLES,
		sendUserMessage: captured.sendUserMessage as unknown as (
			content: string,
			options?: { deliverAs?: "followUp" | "steer" },
		) => Promise<void>,
	});
}

describe("BUILTIN_AGENT_ROLES", () => {
	it("ships writer / reviewer / fact_checker / editor / summarizer", () => {
		expect(Object.keys(BUILTIN_AGENT_ROLES).sort()).toEqual([
			"editor",
			"fact_checker",
			"reviewer",
			"summarizer",
			"writer",
		]);
	});

	it("freezes the registry so hosts cannot mutate built-ins by accident", () => {
		expect(Object.isFrozen(BUILTIN_AGENT_ROLES)).toBe(true);
	});
});

describe("createRequestReviewTool", () => {
	it("declares a tool named request_review with the expected shape", () => {
		const { captured } = makeHarness();
		const tool = makeTool(captured);
		expect(tool.name).toBe("request_review");
		expect(tool.label).toBe("Request Review");
		expect(typeof tool.execute).toBe("function");
	});

	it("dispatches sendUserMessage with the reviewer snippet when no focus is provided", async () => {
		const { captured, ctx } = makeHarness();
		const tool = makeTool(captured);
		const result = await tool.execute(
			"call-1",
			{ role: "reviewer" },
			undefined,
			undefined,
			ctx,
		);
		expect(captured.sendUserMessage).toHaveBeenCalledTimes(1);
		const [content, opts] = captured.sendUserMessage.mock.calls[0]!;
		expect(typeof content).toBe("string");
		expect(content).toContain("[role:reviewer]");
		expect(content).toContain("claim that is not backed by a tool call");
		expect(opts).toEqual({ deliverAs: "followUp" });

		const details = result.details;
		expect(details.role).toBe("reviewer");
		expect(details.resolvedRoleName).toBe("reviewer");
		expect(details.deliveredAs).toBe("followUp");
	});

	it("uses focus to override the role's default scope for one call", async () => {
		const { captured, ctx } = makeHarness();
		const tool = makeTool(captured);
		await tool.execute(
			"call-2",
			{ role: "fact_checker", focus: "Only verify currency amounts" },
			undefined,
			undefined,
			ctx,
		);
		const [content] = captured.sendUserMessage.mock.calls[0]!;
		expect(content).toBe("[role:fact_checker] Only verify currency amounts");
	});

	it("trims whitespace around focus before dispatching", async () => {
		const { captured, ctx } = makeHarness();
		const tool = makeTool(captured);
		await tool.execute(
			"call-2b",
			{ role: "summarizer", focus: "  two sentences max  " },
			undefined,
			undefined,
			ctx,
		);
		const [content] = captured.sendUserMessage.mock.calls[0]!;
		expect(content).toBe("[role:summarizer] two sentences max");
	});

	it("falls back to followUp when deliveredAs is not set", async () => {
		const { captured, ctx } = makeHarness();
		const tool = makeTool(captured);
		await tool.execute("c", { role: "editor" }, undefined, undefined, ctx);
		expect(captured.sendUserMessage.mock.calls[0]![1]).toEqual({ deliverAs: "followUp" });
	});

	it("honours deliveredAs=steer when explicitly configured", async () => {
		const { captured, ctx } = makeHarness();
		const tool = createRequestReviewTool({
			roles: BUILTIN_AGENT_ROLES,
			deliveredAs: "steer",
			sendUserMessage: captured.sendUserMessage as unknown as (
				content: string,
				options?: { deliverAs?: "followUp" | "steer" },
			) => Promise<void>,
		});
		await tool.execute("c", { role: "reviewer" }, undefined, undefined, ctx);
		expect(captured.sendUserMessage.mock.calls[0]![1]).toEqual({ deliverAs: "steer" });
	});

	it("throws with the known-roles list when role is unknown", async () => {
		const { captured, ctx } = makeHarness();
		const tool = makeTool(captured);
		await expect(
			tool.execute("c", { role: "auditor" }, undefined, undefined, ctx),
		).rejects.toThrow(
			/Unknown review role "auditor".*Known roles: editor, fact_checker, reviewer, summarizer, writer/,
		);
	});

	it("text-content result mentions the role and (when present) the focus", async () => {
		const { captured, ctx } = makeHarness();
		const tool = makeTool(captured);
		const withFocus = await tool.execute(
			"c",
			{ role: "reviewer", focus: "claims only" },
			undefined,
			undefined,
			ctx,
		);
		const first = withFocus.content[0];
		expect(first?.type).toBe("text");
		if (first?.type === "text") {
			expect(first.text).toBe("Dispatched reviewer review with focus: claims only.");
		}

		const noFocus = await tool.execute("c2", { role: "summarizer" }, undefined, undefined, ctx);
		const first2 = noFocus.content[0];
		expect(first2?.type).toBe("text");
		if (first2?.type === "text") {
			expect(first2.text).toBe("Dispatched summarizer review.");
		}
	});
});

describe("installAgentTeam", () => {
	it("registers a single request_review tool on the ExtensionAPI", () => {
		const { api, captured } = makeHarness();
		installAgentTeam(api);
		expect(captured.registerTool).toHaveBeenCalledTimes(1);
		const tool = captured.registerTool.mock.calls[0]?.[0] as { name: string };
		expect(tool.name).toBe("request_review");
	});

	it("uses the default roles when no overrides are passed", async () => {
		const { api, captured } = makeHarness();
		installAgentTeam(api);
		const tool = captured.registerTool.mock.calls[0]?.[0] as {
			execute: (...args: never[]) => Promise<unknown>;
		};
		await tool.execute(
			"c" as never,
			{ role: "writer" } as never,
			undefined as never,
			undefined as never,
			{} as never,
		);
		const [content] = captured.sendUserMessage.mock.calls[0]!;
		expect(content).toContain("[role:writer]");
	});

	it("lets hosts add a custom role", async () => {
		const { api, captured } = makeHarness();
		const legal: AgentRole = {
			name: "legal",
			description: "Compliance check.",
			systemPrompt: "Look for compliance issues only.",
		};
		installAgentTeam(api, { additionalRoles: { legal } });
		const tool = captured.registerTool.mock.calls[0]?.[0] as {
			execute: (...args: never[]) => Promise<unknown>;
		};
		await tool.execute(
			"c" as never,
			{ role: "legal" } as never,
			undefined as never,
			undefined as never,
			{} as never,
		);
		const [content] = captured.sendUserMessage.mock.calls[0]!;
		expect(content).toBe("[role:legal] Look for compliance issues only.");
	});

	it("lets hosts override a built-in role", async () => {
		const { api, captured } = makeHarness();
		const strictReviewer: AgentRole = {
			name: "reviewer",
			description: "Strict reviewer.",
			systemPrompt: "Reject unless every claim cites a tool call id.",
		};
		installAgentTeam(api, { additionalRoles: { reviewer: strictReviewer } });
		const tool = captured.registerTool.mock.calls[0]?.[0] as {
			execute: (...args: never[]) => Promise<unknown>;
		};
		await tool.execute(
			"c" as never,
			{ role: "reviewer" } as never,
			undefined as never,
			undefined as never,
			{} as never,
		);
		const [content] = captured.sendUserMessage.mock.calls[0]!;
		expect(content).toBe("[role:reviewer] Reject unless every claim cites a tool call id.");
	});
});

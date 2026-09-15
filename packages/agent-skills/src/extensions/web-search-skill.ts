/**
 * web-search extension — exposes a real `web_search` tool to the agent.
 *
 * Wraps DuckDuckGo's HTML endpoint (no API key required) so the agent has a
 * zero-config search capability. Returns up to `max_results` (default 5) hits
 * with title, snippet, and URL.
 *
 * Failure modes are surfaced as plain text so the model can decide what to do
 * (try a different query, ask the user for clarification, or fall back to
 * its training data). The tool never throws into the agent loop.
 */

import { defineTool, type ExtensionAPI } from "@earendil-works/pi-coding-agent"
import Type from "typebox"

const WebSearchParams = Type.Object({
	query: Type.String({ minLength: 2, maxLength: 500, description: "search query" }),
	max_results: Type.Optional(
		Type.Integer({ minimum: 1, maximum: 10, default: 5, description: "max hits to return" }),
	),
})

type WebSearchArgs = {
	query: string
	max_results?: number
}

export interface SearchHit {
	title: string
	snippet: string
	url: string
}

interface SearchResult {
	hits: SearchHit[]
	source: string
	query: string
	error?: string
}

/** Lightweight HTML entity decoder for the few entities DDG emits. */
function decodeEntities(s: string): string {
	return s
		.replace(/&amp;/g, "&")
		.replace(/&lt;/g, "<")
		.replace(/&gt;/g, ">")
		.replace(/&quot;/g, '"')
		.replace(/&#39;/g, "'")
		.replace(/&nbsp;/g, " ")
}

function createWebSearchTool() {
	return defineTool<typeof WebSearchParams, SearchResult>({
		name: "web_search",
		label: "Web Search",
		description:
			"Search the public web via DuckDuckGo HTML. Use when the user asks about " +
			"current events, public information not in your training data, or to " +
			"verify a fact. Returns title, snippet, and URL for up to 10 hits. " +
			"DO NOT use for tasks that don't need the web — read documents or use " +
			"the office tools first.",
		promptSnippet: "web_search(query, max_results=5) — public web search",
		promptGuidelines: [
			"Keep the query short and specific — three to six words work best.",
			"If the first search returns no useful hits, rephrase and retry.",
			"Always cite the URLs you find when you surface facts from search results.",
		],
		parameters: WebSearchParams,
		async execute(_toolCallId, params: WebSearchArgs, signal) {
			const maxResults = params.max_results ?? 5
			const start = Date.now()
			try {
				const url = `https://html.duckduckgo.com/html/?q=${encodeURIComponent(params.query)}&kl=us-en`
				const response = await fetch(url, {
					signal,
					headers: {
						"User-Agent":
							"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
						"Accept": "text/html",
						"Accept-Language": "en-US,en;q=0.9",
					},
				})
				if (!response.ok) {
					return {
						content: [
							{
								type: "text" as const,
								text: `web_search failed: HTTP ${response.status} ${response.statusText}`,
							},
						],
						details: {
							hits: [],
							source: "duckduckgo",
							query: params.query,
							error: `HTTP ${response.status}`,
						},
					}
				}
				const html = await response.text()
				const hits = parseDuckDuckGo(html, maxResults)
				const summary = hits.length === 0
					? `web_search(${params.query}): no hits from DuckDuckGo.`
					: `web_search(${params.query}) → ${hits.length} hits in ${Date.now() - start}ms:\n\n` +
						hits.map((h, i) => `${i + 1}. ${h.title}\n   ${h.snippet}\n   ${h.url}`).join("\n\n")
				return {
					content: [{ type: "text" as const, text: summary }],
					details: { hits, source: "duckduckgo", query: params.query },
				}
			} catch (err) {
				const message = err instanceof Error ? err.message : String(err)
				return {
					content: [{ type: "text" as const, text: `web_search error: ${message}` }],
					details: { hits: [], source: "duckduckgo", query: params.query, error: message },
				}
			}
		},
	})
}

/**
 * Extract result blocks from DuckDuckGo's HTML. The selectors are best-effort
 * — DDG rotates its markup occasionally. We try a few variants and return
 * what we can parse.
 */
export function parseDuckDuckGo(html: string, maxResults: number): SearchHit[] {
	const hits: SearchHit[] = []

	// Variant 1: <a class="result__a" href="...">title</a> + <a class="result__snippet">snippet</a>
	const blockRe = /<a[^>]*class="result__a"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>[\s\S]*?(?:class="result__snippet"[^>]*>([\s\S]*?)<\/a>|class="result__snippet"[^>]*>([\s\S]*?)<\/td>)/g
	let match: RegExpExecArray | null
	while ((match = blockRe.exec(html)) !== null && hits.length < maxResults) {
		const url = decodeEntities(match[1])
		const title = decodeEntities(stripTags(match[2])).trim()
		const snippet = decodeEntities(stripTags(match[3] || match[4] || "")).trim()
		if (title && url.startsWith("http")) {
			hits.push({ title, snippet, url })
		}
	}

	// Variant 2: <h2 class="result__title"><a href="...">title</a></h2> + <a class="result__url">url</a> + <td class="result__snippet">snippet</td>
	if (hits.length === 0) {
		const titleRe = /<h2[^>]*class="result__title"[^>]*>[\s\S]*?<a[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>[\s\S]*?<\/h2>/g
		const snippetRe = /<a[^>]*class="result__snippet"[^>]*>([\s\S]*?)<\/a>/g
		while ((match = titleRe.exec(html)) !== null && hits.length < maxResults) {
			const url = decodeEntities(match[1])
			const title = decodeEntities(stripTags(match[2])).trim()
			hits.push({ title, snippet: "", url })
		}
		while ((match = snippetRe.exec(html)) !== null && hits.length > 0) {
			hits[hits.length - 1].snippet = decodeEntities(stripTags(match[1])).trim()
		}
	}

	return hits
}

function stripTags(s: string): string {
	return s.replace(/<[^>]+>/g, "").replace(/\s+/g, " ")
}

export function createWebSearchExtension() {
	return (pi: ExtensionAPI): void => {
		pi.registerTool(createWebSearchTool())
	}
}

export const webSearchExtensionDefaults = {
	name: "web-search",
	version: "1.0.0",
	tools: ["web_search"],
	scopes: ["network:out"],
} as const

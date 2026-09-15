/**
 * image-search extension — exposes `image_search` and `fetch_image` tools.
 *
 * Image search hits DuckDuckGo's image endpoint and extracts image URLs from
 * the result markup. `fetch_image` downloads and base64-encodes an image so
 * the model can read it as a multimodal input (or attach it to a doc).
 *
 * Both tools degrade gracefully — no API key required, no auth flow. If DDG
 * blocks the request we surface the error so the agent can ask the user for
 * an alternative URL.
 */

import { defineTool, type ExtensionAPI } from "@earendil-works/pi-coding-agent"
import Type from "typebox"

const ImageSearchParams = Type.Object({
	query: Type.String({ minLength: 2, maxLength: 500 }),
	max_results: Type.Optional(Type.Integer({ minimum: 1, maximum: 8, default: 4 })),
})

const FetchImageParams = Type.Object({
	url: Type.String({ maxLength: 4096 }),
})

const MAX_IMAGE_BYTES = 20 * 1024 * 1024

type ImageSearchArgs = { query: string; max_results?: number }
type FetchImageArgs = { url: string }

interface ImageHit {
	title: string
	thumbnail: string
	page: string
}

interface FetchedImage {
	base64: string
	mime: string
	bytes: number
}

function createImageSearchTool() {
	return defineTool<typeof ImageSearchParams, { hits: ImageHit[] }>({
		name: "image_search",
		label: "Image Search",
		description:
			"Search the public web for images via DuckDuckGo's image endpoint. " +
			"Returns thumbnail URL, source page, and title. Combine with fetch_image " +
			"to read the actual image. Use for: stock photos, icons, diagrams, " +
			"reference imagery. NOT for: copyrighted content you don't have rights to.",
		promptSnippet: "image_search(query, max_results=4) — public image search",
		promptGuidelines: [
			"Use specific queries: 'flat icon rocket' beats 'rocket'.",
			"After image_search, call fetch_image on the page URL to read the image.",
		],
		parameters: ImageSearchParams,
		async execute(_id, params: ImageSearchArgs, signal) {
			const max = params.max_results ?? 4
			try {
				const url = `https://duckduckgo.com/?q=${encodeURIComponent(params.query)}&iax=images&ia=images`
				const response = await fetch(url, {
					signal,
					headers: {
						"User-Agent":
							"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 Safari/605.1.15",
						"Accept": "text/html",
					},
				})
				if (!response.ok) {
					return {
						content: [{ type: "text" as const, text: `image_search failed: HTTP ${response.status}` }],
						details: { hits: [] },
					}
				}
				const html = await response.text()
				const hits = parseDuckDuckGoImages(html, max)
				const summary = hits.length === 0
					? `image_search(${params.query}): no images found. Try rephrasing.`
					: `image_search(${params.query}) → ${hits.length} hits:\n\n` +
						hits.map((h, i) => `${i + 1}. ${h.title}\n   thumb: ${h.thumbnail}\n   page:  ${h.page}`).join("\n\n")
				return {
					content: [{ type: "text" as const, text: summary }],
					details: { hits },
				}
			} catch (err) {
				const message = err instanceof Error ? err.message : String(err)
				return {
					content: [{ type: "text" as const, text: `image_search error: ${message}` }],
					details: { hits: [] },
				}
			}
		},
	})
}

function createFetchImageTool() {
	return defineTool<typeof FetchImageParams, FetchedImage>({
		name: "fetch_image",
		label: "Fetch Image",
		description:
			"Download an image from a URL and return it as base64-encoded bytes " +
			"plus its MIME type. Use after image_search when you want the model " +
			"to actually see the image. Max 20 MB. Returns null content if the " +
			"fetch fails or the size limit is exceeded.",
		promptSnippet: "fetch_image(url) → { base64, mime, bytes }",
		promptGuidelines: [
			"Prefer HTTPS URLs (DuckDuckGo thumbnails redirect to https).",
			"If fetch_image returns null, the URL may be rate-limited or too large.",
		],
		parameters: FetchImageParams,
		async execute(_id, params: FetchImageArgs, signal) {
			if (!params.url.startsWith("http")) {
				return {
					content: [{ type: "text" as const, text: "fetch_image: URL must be http(s)" }],
					details: { base64: "", mime: "", bytes: 0 },
				}
			}
			try {
				const response = await fetch(params.url, {
					signal,
					headers: {
						"User-Agent":
							"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 Safari/605.1.15",
					},
				})
				if (!response.ok || !response.body) {
					return {
						content: [{ type: "text" as const, text: `fetch_image: HTTP ${response.status}` }],
						details: { base64: "", mime: "", bytes: 0 },
					}
				}
				const reader = response.body.getReader()
				const chunks: Buffer[] = []
				let total = 0
				for (;;) {
					const part = await reader.read()
					if (part.done) break
					total += part.value.byteLength
					if (total > MAX_IMAGE_BYTES) {
						await reader.cancel()
						return {
							content: [{ type: "text" as const, text: `fetch_image: exceeds ${MAX_IMAGE_BYTES} bytes` }],
							details: { base64: "", mime: "", bytes: 0 },
						}
					}
					chunks.push(Buffer.from(part.value))
				}
				const contentType = response.headers.get("content-type") ?? ""
				const mime = contentType.includes("png")
					? "image/png"
					: contentType.includes("gif")
						? "image/gif"
						: contentType.includes("webp")
							? "image/webp"
							: "image/jpeg"
				const base64 = Buffer.concat(chunks).toString("base64")
				return {
					content: [
						{
							type: "text" as const,
							text: `fetch_image: ${total} bytes, ${mime}. base64 length ${base64.length}`,
						},
					],
					details: { base64, mime, bytes: total },
				}
			} catch (err) {
				const message = err instanceof Error ? err.message : String(err)
				return {
					content: [{ type: "text" as const, text: `fetch_image error: ${message}` }],
					details: { base64: "", mime: "", bytes: 0 },
				}
			}
		},
	})
}

/**
 * Extract image result JSON from DDG. The endpoint embeds a JSON blob in a
 * script tag — much more reliable than scraping markup. We try parsing from
 * every candidate `{` and accept the first one that yields an `image_results`
 * or `results` array; this handles nested objects that defeat a non-greedy
 * regex.
 */
export function parseDuckDuckGoImages(html: string, maxResults: number): ImageHit[] {
	const hits: ImageHit[] = []
	// Walk every `DDG.pageLayout.DDG.result = ` anchor; for each, find the
	// first balanced JSON object after the `=` and try parsing it.
	const anchorRe = /DDG\.pageLayout\.DDG\.result\s*=\s*/g
	let anchor: RegExpExecArray | null
	while ((anchor = anchorRe.exec(html)) !== null) {
		const start = anchor.index + anchor[0].length
		const end = findBalancedJsonEnd(html, start)
		if (end === -1) continue
		const candidate = html.slice(start, end)
		try {
			const blob = JSON.parse(candidate)
			const images = blob?.["image_results"] ?? blob?.results ?? []
			if (Array.isArray(images) && images.length > 0) {
				for (const item of images.slice(0, maxResults)) {
					hits.push({
						title: typeof item.title === "string" ? item.title : "",
						thumbnail: typeof item.thumbnail === "string"
							? item.thumbnail
							: typeof item.image === "string"
								? item.image
								: "",
						page: typeof item.url === "string" ? item.url : "",
					})
				}
				return hits
			}
		} catch {
			// try the next `{`
		}
	}
	// Fallback: regex on the visible snippet
	const re = /<a[^>]*class="result__a"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>[\s\S]*?<img[^>]*src="([^"]+)"/g
	let m: RegExpExecArray | null
	while ((m = re.exec(html)) !== null && hits.length < maxResults) {
		hits.push({ title: m[2].replace(/<[^>]+>/g, "").trim(), page: m[1], thumbnail: m[3] })
	}
	return hits
}

/**
 * Walk from `start` to the matching closing `}` honouring string literals and
 * escapes. Returns -1 if no balanced close is found.
 */
function findBalancedJsonEnd(s: string, start: number): number {
	let depth = 0
	let inString = false
	let escape = false
	for (let i = start; i < s.length; i++) {
		const ch = s[i]
		if (escape) {
			escape = false
			continue
		}
		if (inString) {
			if (ch === "\\") {
				escape = true
			} else if (ch === '"') {
				inString = false
			}
			continue
		}
		if (ch === '"') {
			inString = true
		} else if (ch === "{") {
			depth++
		} else if (ch === "}") {
			depth--
			if (depth === 0) return i + 1
		}
	}
	return -1
}



export function createImageSearchExtension() {
	return (pi: ExtensionAPI): void => {
		pi.registerTool(createImageSearchTool())
		pi.registerTool(createFetchImageTool())
	}
}

export const imageSearchExtensionDefaults = {
	name: "image-search",
	version: "1.0.0",
	tools: ["image_search", "fetch_image"],
	scopes: ["network:out"],
} as const

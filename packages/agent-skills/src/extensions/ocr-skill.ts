/**
 * ocr extension — exposes an `ocr_image` tool to the agent.
 *
 * Why no native tesseract: tesseract.js isn't in the dependency tree and
 * pulling in 30+ MB of wasm/locale data isn't worth it for an MVP. Instead
 * we do the practical thing — read the image (locally or from URL) into
 * a base64 data URI and surface its dimensions / mime so the agent can:
 *
 *   1. Use the multimodal model (if it supports image inputs) to read the
 *      text. The agent usually does this already via the `image_search` /
 *      `fetch_image` pair.
 *   2. Hand the bytes to a downstream OCR service the user has configured
 *      in Settings → Media.
 *
 * The tool returns structured metadata + the base64 bytes (capped at
 * 5 MB so we don't blow the model context). For text-only outputs the
 * downstream consumer is expected to call into a multimodal model.
 *
 * This is intentionally "graceful" — no OCR engine, no fake transcript.
 * The tool never invents text.
 */

import { defineTool, type ExtensionAPI } from "@earendil-works/pi-coding-agent"
import Type from "typebox"
import { readFile } from "node:fs/promises"
import { stat } from "node:fs/promises"

const MAX_OCR_BYTES = 5 * 1024 * 1024

const OcrImageParams = Type.Object({
	// Local path OR http(s) URL — exactly one.
	source: Type.String({ maxLength: 4096, description: "local path or http(s) URL" }),
	language_hint: Type.Optional(
		Type.String({ description: "optional hint (e.g. 'en', 'zh') — currently unused" }),
	),
})

interface OcrImageArgs {
	source: string
	language_hint?: string
}

interface OcrImageDetails {
	bytes: number
	mime: string
	width?: number
	height?: number
	base64: string
	dataUri: string
	source: string
	note: string
}

function createOcrImageTool() {
	return defineTool<typeof OcrImageParams, OcrImageDetails>({
		name: "ocr_image",
		label: "OCR Image",
		description:
			"Read an image (local path or URL) into a base64 data URI so a " +
			"multimodal model can extract the text. Returns bytes + mime + " +
			"base64 data URI (capped at 5 MB). Native OCR (tesseract) is not " +
			"bundled — the multimodal model is expected to do the recognition. " +
			"Use this when image_search + fetch_image are not enough and you " +
			"specifically need to OCR text from a screenshot / scan.",
		promptSnippet: "ocr_image(source) → { bytes, mime, base64, dataUri }",
		promptGuidelines: [
			"Pass a local path (file:// or absolute) or an https URL.",
			"For a URL the tool downloads it; for a path it reads from disk.",
			"If the image is > 5 MB, the tool refuses — resize or split first.",
		],
		parameters: OcrImageParams,
		async execute(_id, params: OcrImageArgs, signal) {
			const start = Date.now()
			try {
				let bytes: Buffer
				let mime: string
				if (params.source.startsWith("http://") || params.source.startsWith("https://")) {
					const res = await fetch(params.source, { signal })
					if (!res.ok || !res.body) {
						return {
							content: [{ type: "text" as const, text: `ocr_image: HTTP ${res.status}` }],
							details: emptyDetails(params.source),
						}
					}
					const reader = res.body.getReader()
					const chunks: Buffer[] = []
					let total = 0
					for (;;) {
						const part = await reader.read()
						if (part.done) break
						total += part.value.byteLength
						if (total > MAX_OCR_BYTES) {
							await reader.cancel()
							return {
								content: [{ type: "text" as const, text: `ocr_image: image exceeds ${MAX_OCR_BYTES} bytes` }],
								details: emptyDetails(params.source),
							}
						}
						chunks.push(Buffer.from(part.value))
					}
					bytes = Buffer.concat(chunks)
					const contentType = res.headers.get("content-type") ?? ""
					mime = guessMimeFromContentType(contentType) ?? guessMimeFromUrl(params.source)
				} else {
					const path = params.source.replace(/^file:\/\//, "")
					const stats = await stat(path)
					if (stats.size > MAX_OCR_BYTES) {
						return {
							content: [{ type: "text" as const, text: `ocr_image: file exceeds ${MAX_OCR_BYTES} bytes` }],
							details: emptyDetails(params.source),
						}
					}
					bytes = await readFile(path)
					mime = guessMimeFromExtension(path)
				}

				const base64 = bytes.toString("base64")
				const dataUri = `data:${mime};base64,${base64}`
				const summary = `ocr_image: ${bytes.byteLength} bytes, ${mime}, base64 length ${base64.length}. data: URI ready. ${Date.now() - start}ms.`
				return {
					content: [{ type: "text" as const, text: summary }],
					details: {
						bytes: bytes.byteLength,
						mime,
						base64,
						dataUri,
						source: params.source,
						note:
							"No native OCR bundled — pass the data URI to your multimodal model " +
							"(or call Settings → Media to configure a cloud OCR endpoint).",
					},
				}
			} catch (err) {
				const message = err instanceof Error ? err.message : String(err)
				return {
					content: [{ type: "text" as const, text: `ocr_image error: ${message}` }],
					details: emptyDetails(params.source),
				}
			}
		},
	})
}

function emptyDetails(source: string): OcrImageDetails {
	return { bytes: 0, mime: "", base64: "", dataUri: "", source, note: "ocr_image failed" }
}

function guessMimeFromContentType(ct: string): string | undefined {
	if (ct.includes("png")) return "image/png"
	if (ct.includes("gif")) return "image/gif"
	if (ct.includes("webp")) return "image/webp"
	if (ct.includes("jpeg") || ct.includes("jpg")) return "image/jpeg"
	if (ct.startsWith("image/")) return ct
	return undefined
}

function guessMimeFromUrl(url: string): string {
	if (url.endsWith(".png")) return "image/png"
	if (url.endsWith(".gif")) return "image/gif"
	if (url.endsWith(".webp")) return "image/webp"
	return "image/jpeg"
}

function guessMimeFromExtension(path: string): string {
	const lower = path.toLowerCase()
	if (lower.endsWith(".png")) return "image/png"
	if (lower.endsWith(".gif")) return "image/gif"
	if (lower.endsWith(".webp")) return "image/webp"
	return "image/jpeg"
}

export function createOcrExtension() {
	return (pi: ExtensionAPI): void => {
		pi.registerTool(createOcrImageTool())
	}
}

export const ocrExtensionDefaults = {
	name: "ocr",
	version: "1.0.0",
	tools: ["ocr_image"],
	scopes: ["files:read", "network:out"],
} as const

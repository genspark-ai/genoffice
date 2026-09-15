import { describe, expect, it } from "vitest"
import { createOcrExtension, ocrExtensionDefaults } from "../src/extensions/ocr-skill"

interface RegisteredTool {
	name: string
	label: string
	description: string
	parameters: unknown
	execute: (
		toolCallId: string,
		params: unknown,
		signal: AbortSignal | undefined,
	) => Promise<{ content: Array<{ type: string; text: string }>; details: unknown }>
}

function makeFakePi() {
	const tools = new Map<string, RegisteredTool>()
	return {
		tools,
		registerTool(tool: RegisteredTool) {
			tools.set(tool.name, tool)
		},
	}
}

describe("ocr-skill", () => {
	it("exports stable defaults", () => {
		expect(ocrExtensionDefaults.tools).toEqual(["ocr_image"])
		expect(ocrExtensionDefaults.scopes).toContain("files:read")
	})

	it("registers an ocr_image tool", () => {
		const pi = makeFakePi()
		createOcrExtension()(pi as never)
		expect(pi.tools.has("ocr_image")).toBe(true)
		const tool = pi.tools.get("ocr_image")!
		expect(tool.label).toBe("OCR Image")
		expect(tool.description).toMatch(/multimodal/i)
	})

	it("ocr_image returns graceful envelope on http failure", async () => {
		const pi = makeFakePi()
		createOcrExtension()(pi as never)
		const tool = pi.tools.get("ocr_image")!
		const controller = new AbortController()
		controller.abort()
		const result = await tool.execute(
			"call_1",
			{ source: "https://example.com/missing.png" },
			controller.signal,
		)
		expect(result.content[0].type).toBe("text")
		expect(result.details).toBeDefined()
	})

	it("ocr_image refuses non-http non-path sources", async () => {
		const pi = makeFakePi()
		createOcrExtension()(pi as never)
		const tool = pi.tools.get("ocr_image")!
		const result = await tool.execute(
			"call_1",
			{ source: "ftp://bad" },
			undefined,
		)
		expect(result.content[0].text).toMatch(/ocr_image/)
		expect((result.details as { bytes: number }).bytes).toBe(0)
	})
})

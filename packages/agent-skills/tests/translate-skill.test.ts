import { afterEach, beforeEach, describe, expect, it } from "vitest"
import {
  createTranslateSkillExtension,
  ALL_TRANSLATE_TOOL_NAMES,
  __setReadSettingsForTests,
  __setTranslateOneForTests,
  __setChatForProviderForTests,
  __resetKbForTests,
} from "../src/extensions/translate-skill"
import type { AiSettings } from "@genoffice/ai-provider"

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

function fakeSettings(overrides: Partial<AiSettings> = {}): AiSettings {
  return {
    provider: "genspark",
    providers: {
      genspark: { apiKey: "test-key", model: "gpt-4o-mini" },
    } as AiSettings["providers"],
    ...overrides,
  } as AiSettings
}

describe("translate-skill", () => {
  afterEach(() => {
    __setReadSettingsForTests(null)
    __setTranslateOneForTests(null)
    __setChatForProviderForTests(null)
    __resetKbForTests()
  })

  it("exposes the 6-tool surface in ALL_TRANSLATE_TOOL_NAMES", () => {
    expect(ALL_TRANSLATE_TOOL_NAMES).toEqual([
      "translate_text",
      "translate_file",
      "build_dictionary",
      "kb_search",
      "kb_upsert",
      "kb_remove",
    ])
  })

  it("registers all 6 tools on the pi extension", () => {
    const pi = makeFakePi()
    createTranslateSkillExtension()(pi as never)
    for (const name of ALL_TRANSLATE_TOOL_NAMES) {
      expect(pi.tools.has(name), `missing tool ${name}`).toBe(true)
    }
  })

  it("translate_text returns the upstream TranslateResponse verbatim", async () => {
    const pi = makeFakePi()
    __setReadSettingsForTests(async () => fakeSettings())
    __setTranslateOneForTests(async () => ({
      ok: true,
      translated: "你好,世界",
      status: "translated",
      matchedTerms: ["hello"],
      warnings: [],
    }) as never)

    createTranslateSkillExtension()(pi as never)
    const tool = pi.tools.get("translate_text")!
    const out = await tool.execute("c1", {
      text: "hello world",
      target_lang: "zh-CN",
    }, undefined)

    expect(out.details).toMatchObject({
      ok: true,
      translated: "你好,世界",
      status: "translated",
      matchedTerms: ["hello"],
    })
    expect(out.content[0].text).toContain("translate_text → status=translated")
  })

  it("translate_text surfaces failures with a stable error envelope", async () => {
    const pi = makeFakePi()
    __setReadSettingsForTests(async () => fakeSettings())
    __setTranslateOneForTests(async () => ({ ok: false, error: "no api key" }) as never)

    createTranslateSkillExtension()(pi as never)
    const out = await pi.tools.get("translate_text")!.execute(
      "c1",
      { text: "hello", target_lang: "zh-CN" },
      undefined,
    )
    expect(out.details).toMatchObject({ ok: false, error: "no api key" })
    expect(out.content[0].text).toMatch(/failed/i)
  })

  it("build_dictionary asks the LLM, parses pairs, and writes a JSON file", async () => {
    const pi = makeFakePi()
    __setReadSettingsForTests(async () => fakeSettings())
    __setChatForProviderForTests(async () => ({
      ok: true,
      content: "SKUA : 面料 A\nfabric code : 面料编号\nbrandX : 品牌X\n",
    }) as never)

    createTranslateSkillExtension()(pi as never)
    const tool = pi.tools.get("build_dictionary")!
    const outPath = `${process.env.TMPDIR ?? "/tmp"}/translate-skill-test-${Date.now()}.dictionary.json`
    const out = await tool.execute(
      "c1",
      { input_path: "/tmp/fake.pdf", target_lang: "zh-CN", output_path: outPath, max_pairs: 10 },
      undefined,
    )
    const details = out.details as { ok: boolean; pairCount: number; outputPath: string }
    expect(details.ok).toBe(true)
    expect(details.pairCount).toBe(3)
    expect(details.outputPath).toBe(outPath)
  })
})

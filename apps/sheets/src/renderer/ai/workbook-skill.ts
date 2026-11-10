import type { AgentSkill } from '@genoffice/agent-core'
import basePrompt from './prompts/base.md?raw'
import { verifySheetsResponse } from './response-verify'
import { placeholderInstruction } from './redact'
import {
  WORKBOOK_TOOLS,
  buildWorkbookContext,
  executeWorkbookTool,
  type SheetsSkillDeps,
} from './tools'

/**
 * The workbook DSL as an AgentSkill: mirrors createDocsSkill's shape
 * (systemPrompt + tools + buildContext + executeTool) so it plugs into the
 * same packages/agent-core AgentLoop docx uses.
 *
 * Prompt layout: the always-loaded base prompt (prompts/base.md) stays small
 * — workflow, op catalog, cross-cutting discipline — while per-domain field
 * definitions and conventions live in prompts/guides/*.md, loaded on demand
 * via load_guide.
 *
 * The base prompt is imported raw and therefore fixed at build time, so the
 * placeholder section is a getter rather than a string: which labels exist is
 * a property of the open workbook, and a model told about a placeholder that
 * is not there — or not told about one that is — misreads the sheet either way.
 */
export function createWorkbookSkill(deps: SheetsSkillDeps): AgentSkill {
  return {
    id: 'sheets',
    get systemPrompt() {
      const index = deps.redactions?.()
      const labels = index?.isEmpty
        ? []
        : (deps.getActiveSheetInfo().sheets.flatMap((sheet) => index?.labelsFor(sheet.id) ?? []) ??
          [])
      return labels.length > 0 ? `${basePrompt}\n\n${placeholderInstruction(labels)}` : basePrompt
    },
    tools: WORKBOOK_TOOLS,
    buildContext: () => buildWorkbookContext(deps),
    executeTool: (call) => executeWorkbookTool(call, deps),
    verifyResponse: verifySheetsResponse,
  }
}

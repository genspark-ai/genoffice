import type { AgentSkill } from '@genoffice/agent-core'
import {
  READ_ATTACHMENT_TOOL_DEF,
  READ_KNOWLEDGE_FILE_TOOL,
  SEARCH_KNOWLEDGE_BASE_TOOL,
} from '@genoffice/agent-core/resource-tools'
import { DOCS_OFFICE_CLI_TOOL } from './cli-tool-schema'

/** Host-authorized attachment, knowledge and bounded Office helpers. No model service. */
export type DocsResourceTools = Pick<AgentSkill, 'tools' | 'buildContext' | 'executeTool'>

/** Contracts for resource tools that a host may authorize for the current document. */
export const DOCS_RESOURCE_TOOLS = [
  READ_ATTACHMENT_TOOL_DEF,
  READ_KNOWLEDGE_FILE_TOOL,
  SEARCH_KNOWLEDGE_BASE_TOOL,
  DOCS_OFFICE_CLI_TOOL,
]
export const DOCS_RESOURCE_TOOL_NAMES = new Set(DOCS_RESOURCE_TOOLS.map((tool) => tool.name))

/** File-only commands. Provider configuration and all inference stay in the Harness. */
export const DOCS_FILE_COMMANDS = new Set([
  'info',
  'convert',
  'create',
  'slides',
  'sheet',
  'docs',
  'merge',
  'pdf',
  'render',
  'guide',
  '--help',
])

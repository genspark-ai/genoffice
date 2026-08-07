export type {
  AgentImage,
  AgentMessage,
  AgentStreamCallbacks,
  AgentStreamHandle,
  AgentStreamRequest,
  AgentToolCall,
  AgentToolDef,
  AgentToolResult,
  AgentTransport,
  ToolDisplay,
  ToolExecution,
} from './types'
export { composeSkills } from './skill'
export type { AgentSkill } from './skill'
export {
  INSTRUCTION_SCOPES,
  LOAD_SKILL_TOOL,
  buildInstructionsPrompt,
  buildRulesPrompt,
  buildSkillsPrompt,
  coerceScope,
  isInstructionScope,
  parseSkillMarkdown,
  scopeApplies,
  serializeSkillMarkdown,
  skillBodyForTool,
  skillsForSurface,
} from './instructions'
export type {
  AgentRules,
  AppSurface,
  InstructionScope,
  UserSkill,
  UserSkillSummary,
} from './instructions'
export { createWebSkill } from './web-skill'
export type {
  BrowsePageBridgeResult,
  ExtractPagesBridgeResult,
  WebSkillBridge,
  WebSkillOptions,
} from './web-skill'
export { AgentLoop, COMPLETED_VIA_TOOLS_TEXT, sanitizeAgentPayload } from './loop'
export type {
  AgentLoopEvents,
  AgentLoopOptions,
  AgentRunResult,
  CompactionOptions,
  ToolExecutedEvent,
} from './loop'
export { createIpcTransport, IPC_STREAM_SILENCE_TIMEOUT_MS } from './electron-transport'
export type { IpcStreamChunk, IpcStreamStart, IpcTransportOptions } from './electron-transport'

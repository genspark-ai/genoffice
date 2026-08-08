/**
 * Contract for the standalone settings window (renderer/settings.html).
 *
 * Kept self-contained, like home-api.ts: the preload is bundled as a single
 * file and must not pull in a shared runtime chunk, so these types mirror the
 * ones in @genoffice/ai-provider and @genoffice/agent-core rather than
 * importing them.
 */

/** sections in the sidebar, in display order */
export const SETTINGS_SECTIONS = ['model', 'network', 'rules', 'skills', 'general'] as const
export type SettingsSection = (typeof SETTINGS_SECTIONS)[number]

/** mirrors ReasoningEffort in @genoffice/ai-provider */
export type ReasoningEffort = 'minimal' | 'low' | 'medium' | 'high'

/** mirrors InstructionScope in @genoffice/agent-core */
export type InstructionScope = 'global' | 'docx' | 'pptx' | 'sheets' | 'pdf'

export const INSTRUCTION_SCOPES: readonly InstructionScope[] = [
  'global',
  'docx',
  'pptx',
  'sheets',
  'pdf',
]

/** mirrors AiProfileSummary in @genoffice/ai-provider */
export interface AiProfileSummary {
  id: string
  label: string
  baseUrl: string
  model: string
  apiKey: string
}

/** mirrors AiModelSettings in @genoffice/ai-provider */
export interface AiModelSettings {
  mode: 'genspark' | 'custom'
  /** saved endpoints; the fields below belong to `profileId` */
  profiles: AiProfileSummary[]
  profileId: string | null
  baseUrl: string
  model: string
  apiKey: string
  temperature: number | null
  maxTokens: number | null
  reasoningEffort: ReasoningEffort | null
  tavilyApiKey: string
  proxyUrl: string
}

export interface AiProviderTestResult {
  ok: boolean
  error?: string
}

/** mirrors AgentRules in @genoffice/agent-core */
export type AgentRules = Partial<Record<InstructionScope, string>>

/** mirrors UserMemory in @genoffice/agent-core */
export interface UserMemory {
  id: string
  text: string
  createdAt: number
}

/** mirrors UserSkill in @genoffice/agent-core */
export interface UserSkill {
  id: string
  name: string
  description: string
  scopes: InstructionScope[]
  body: string
  enabled: boolean
}

export interface SettingsSnapshot {
  ai: AiModelSettings
  rules: AgentRules
  skills: UserSkill[]
  /** what the agent recorded about the user; shown so it can be corrected */
  memories: UserMemory[]
  language: string
  updateChannel: 'stable' | 'beta'
  appVersion: string
  account: { loggedIn: boolean; email?: string }
}

export interface SettingsApi {
  /** everything the window renders, in one round trip */
  load(): Promise<SettingsSnapshot>
  saveAi(settings: AiModelSettings): Promise<AiModelSettings>
  testProvider(settings: Omit<AiModelSettings, 'mode'>): Promise<AiProviderTestResult>
  saveRules(rules: AgentRules): Promise<AgentRules>
  saveSkill(skill: Partial<UserSkill> & { name: string; body: string }): Promise<UserSkill>
  deleteSkill(id: string): Promise<void>
  deleteMemory(id: string): Promise<void>
  /** native multi-select picker; returns the skills that were imported */
  importSkillFiles(): Promise<UserSkill[]>
  /** drag-and-drop path: the renderer already read the files */
  importSkillContents(files: Array<{ filename: string; content: string }>): Promise<UserSkill[]>
  setLanguage(lang: string): Promise<void>
  setUpdateChannel(channel: 'stable' | 'beta'): Promise<void>
  accountLogin(): Promise<boolean>
  accountLogout(): Promise<void>
  close(): Promise<void>
}

export const SETTINGS_CHANNELS = {
  load: 'settings:load',
  saveAi: 'settings:save-ai',
  testProvider: 'ai:test-provider',
  saveRules: 'settings:save-rules',
  saveSkill: 'settings:save-skill',
  deleteSkill: 'settings:delete-skill',
  deleteMemory: 'settings:delete-memory',
  importSkillFiles: 'settings:import-skill-files',
  importSkillContents: 'settings:import-skill-contents',
  setLanguage: 'settings:set-language',
  setUpdateChannel: 'settings:set-update-channel',
  accountLogin: 'settings:account-login',
  accountLogout: 'settings:account-logout',
  close: 'settings:close',
  /** main → window, after a login completes in the browser */
  changed: 'settings:changed',
} as const

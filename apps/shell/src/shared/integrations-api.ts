import type { InstallOutcome } from '@genoffice/cli/install'

export type AgentId =
  'claude-code' | 'codex' | 'cursor' | 'gemini' | 'copilot' | 'opencode' | 'windsurf'

export interface AgentTarget {
  id: AgentId
  /** product name, shown as is (not translated) */
  label: string
  /** global skills directory the agent scans */
  skillsDir: string
}

export type SkillInstallStatus =
  /** no genoffice/ folder in the skills directory */
  | 'missing'
  /** written by this app, same version and bytes as the bundled skill */
  | 'installed'
  /** written by this app, older than the bundled skill */
  | 'outdated'
  /** written by this app, edited since (bytes differ from what was written) */
  | 'modified'
  /** our skill by front matter, but not written by this app (other host, npx, by hand) */
  | 'foreign'
  /** a newer skill version than the one bundled, whoever wrote it */
  | 'newer'
  /** genoffice/ exists but does not hold our skill */
  | 'occupied'

export interface SkillInstallState {
  status: SkillInstallStatus
  /** `<skillsDir>/genoffice/SKILL.md` */
  path: string
  installedVersion?: string
  /** installed version is older than the bundled one (foreign rows offer an update then) */
  older?: boolean
}

export interface CliStatus extends InstallOutcome {
  /** directory holding the genoffice launcher (what `~/.genoffice/launcher` points at) */
  launcherDir: string
  /** app runs from a dmg / AppImage mount: the launcher path will not survive a restart */
  ephemeral: boolean
  /** version of the bundled command line */
  version: string
}

export interface IntegrationsStatus {
  cli: CliStatus
  /** version of the bundled skill (skills/genoffice/SKILL.md front matter) */
  skillVersion: string
  /** lowest CLI version that skill describes */
  skillNeedsCli: string
  agents: Array<AgentTarget & { state: SkillInstallState }>
}

export interface IntegrationsApi {
  /** probe the CLI link and every detected agent's skills directory (no writes) */
  status(): Promise<IntegrationsStatus>
  /** write the bundled skill into one agent's skills directory (or a picked folder) */
  installSkill(target: { agentId: AgentId } | { dir: string }): Promise<SkillInstallState>
  /** remove a skill this app wrote */
  uninstallSkill(agentId: AgentId): Promise<SkillInstallState>
  /** folder picker for "install elsewhere"; null when cancelled */
  pickSkillDir(title: string): Promise<string | null>
  /** save dialog + write of genoffice-skill-<version>.zip; the saved path, null when cancelled */
  saveSkillZip(title: string): Promise<string | null>
  /** put text on the clipboard (paths, the manual PATH command) */
  copyText(text: string): Promise<void>
}

export const INTEGRATIONS_CHANNELS = {
  status: 'integrations:status',
  installSkill: 'integrations:install-skill',
  uninstallSkill: 'integrations:uninstall-skill',
  pickSkillDir: 'integrations:pick-skill-dir',
  saveSkillZip: 'integrations:save-skill-zip',
  copyText: 'integrations:copy-text',
} as const

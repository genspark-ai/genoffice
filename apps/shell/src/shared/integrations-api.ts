import type {
  AgentId,
  AgentTarget,
  SkillInstallState,
  SkillInstallStatus,
} from '@genoffice/cli/agent-skills'
import type { InstallOutcome } from '@genoffice/cli/install'
import type { FoundSkill, RecommendedSkill } from './found-skill'

export type { AgentId, AgentTarget, SkillInstallState, SkillInstallStatus }

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
  /** the skills in GenOffice's own folder; nothing else is scanned */
  listSkills(): Promise<FoundSkill[]>
  /** the body of one listed SKILL.md, frontmatter stripped */
  skillBody(path: string): Promise<string>
  /** open-dialog for a skill folder, then copy it in; null when cancelled */
  importSkillFolder(title: string): Promise<FoundSkill | null>
  /** open-dialog for a skill zip or a single SKILL.md, then bring it in; null when cancelled */
  importSkillZip(title: string): Promise<FoundSkill | null>
  /** the skills GenOffice ships as suggestions, with each one's install state */
  listRecommendedSkills(): Promise<RecommendedSkill[]>
  /** copy (or, when outdated, replace) one recommended skill into our folder */
  installRecommendedSkill(name: string): Promise<FoundSkill>
  /** delete one of our skills by its folder name */
  removeSkill(name: string): Promise<void>
}

export const INTEGRATIONS_CHANNELS = {
  status: 'integrations:status',
  listSkills: 'integrations:list-skills',
  skillBody: 'integrations:skill-body',
  importSkillFolder: 'integrations:import-skill-folder',
  importSkillZip: 'integrations:import-skill-zip',
  listRecommendedSkills: 'integrations:list-recommended-skills',
  installRecommendedSkill: 'integrations:install-recommended-skill',
  removeSkill: 'integrations:remove-skill',
  installSkill: 'integrations:install-skill',
  uninstallSkill: 'integrations:uninstall-skill',
  pickSkillDir: 'integrations:pick-skill-dir',
  saveSkillZip: 'integrations:save-skill-zip',
  copyText: 'integrations:copy-text',
} as const

import type {
  AgentId,
  AgentTarget,
  SkillInstallState,
  SkillInstallStatus,
} from '@genoffice/cli/agent-skills'
import type { InstallOutcome } from '@genoffice/cli/install'
import type { FoundSkill } from './found-skill'

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
  /**
   * Every SKILL.md this machine can offer: the ones under the save directory
   * (usable without asking) first, then the coding agents' own as candidates.
   * Nothing here has been copied or run.
   */
  listSkills(): Promise<FoundSkill[]>
  /** the body of one SKILL.md, frontmatter stripped, for the turn that runs it */
  skillBody(path: string): Promise<string>
  /**
   * Copy a skill the user picked into our own folder. Only ever called from an
   * explicit click — scanning never copies.
   */
  importSkill(path: string): Promise<FoundSkill>
}

export const INTEGRATIONS_CHANNELS = {
  status: 'integrations:status',
  listSkills: 'integrations:list-skills',
  skillBody: 'integrations:skill-body',
  importSkill: 'integrations:import-skill',
  installSkill: 'integrations:install-skill',
  uninstallSkill: 'integrations:uninstall-skill',
  pickSkillDir: 'integrations:pick-skill-dir',
  saveSkillZip: 'integrations:save-skill-zip',
  copyText: 'integrations:copy-text',
} as const

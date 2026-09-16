import { readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { app, clipboard, dialog, ipcMain, type BrowserWindow } from 'electron'
import {
  agentTarget,
  bundledSkillFrom,
  buildSkillZip,
  detectAgents,
  installSkill,
  LEDGER_KEY,
  ledgerFromSettings,
  readInstallState,
  uninstallSkill,
  type BundledSkill,
  type SkillLedger,
} from '@genoffice/cli/agent-skills'
import { inspectCliLink } from '@genoffice/cli/install'
import { readAppSettings, writeAppSetting } from './app-settings'
import { isEphemeralInstall } from './cli-link'
import {
  INTEGRATIONS_CHANNELS,
  type AgentId,
  type IntegrationsStatus,
  type SkillInstallState,
} from '../shared/integrations-api'

export interface IntegrationsDeps {
  settingsPath: () => string
  /** the shell window dialogs attach to */
  window: () => BrowserWindow | null
  /** directory holding genoffice / genoffice.cmd and, packaged, skills/genoffice/SKILL.md */
  cliDir: string
  /** skills/genoffice/SKILL.md (repo file in dev, Resources/cli/skills/... packaged) */
  skillPath: string
  /** packages/cli/package.json (its version is the CLI version) */
  cliPackageJson: string
}

/** Settings → Integrations: probe, install, uninstall, zip. No write happens without a click in that pane. */
export function registerIntegrationsIpc(deps: IntegrationsDeps): void {
  const bundled = (): BundledSkill => bundledSkillFrom(readFileSync(deps.skillPath))
  const ledger = (): SkillLedger => ledgerFromSettings(readAppSettings(deps.settingsPath()))
  const saveLedger = (l: SkillLedger) => writeAppSetting(deps.settingsPath(), LEDGER_KEY, l)
  const stateOf = (skillsDir: string): SkillInstallState =>
    readInstallState(skillsDir, bundled(), ledger())

  ipcMain.handle(INTEGRATIONS_CHANNELS.status, (): IntegrationsStatus => {
    const skill = bundled()
    const l = ledger()
    const launcher = join(deps.cliDir, process.platform === 'win32' ? 'genoffice.cmd' : 'genoffice')
    return {
      cli: {
        ...inspectCliLink({ launcher }),
        launcherDir: deps.cliDir,
        ephemeral: app.isPackaged && isEphemeralInstall(process.resourcesPath, process.env),
        version: cliVersion(deps.cliPackageJson),
      },
      skillVersion: skill.version,
      skillNeedsCli: /^\s+cli:\s*['"]?>=\s*(\d+\.\d+\.\d+)/m.exec(skill.text)?.[1] ?? '',
      agents: detectAgents().map((a) => ({ ...a, state: readInstallState(a.skillsDir, skill, l) })),
    }
  })

  ipcMain.handle(
    INTEGRATIONS_CHANNELS.installSkill,
    (_e, target: { agentId?: AgentId; dir?: string }): SkillInstallState => {
      const skillsDir = target.dir ?? agentTarget(target.agentId!)?.skillsDir
      if (!skillsDir) throw new Error('unknown skill target')
      const l = ledger()
      installSkill(skillsDir, bundled(), l)
      saveLedger(l)
      return stateOf(skillsDir)
    },
  )

  ipcMain.handle(
    INTEGRATIONS_CHANNELS.uninstallSkill,
    (_e, agentId: AgentId): SkillInstallState => {
      const target = agentTarget(agentId)
      if (!target) throw new Error('unknown skill target')
      const l = ledger()
      if (uninstallSkill(target.skillsDir, l)) saveLedger(l)
      return stateOf(target.skillsDir)
    },
  )

  // dialog titles come from the renderer, which owns the UI language
  ipcMain.handle(
    INTEGRATIONS_CHANNELS.pickSkillDir,
    async (_e, title: string): Promise<string | null> => {
      const opts: Electron.OpenDialogOptions = {
        title: String(title ?? ''),
        properties: ['openDirectory', 'createDirectory'],
      }
      const win = deps.window()
      const r = await (win ? dialog.showOpenDialog(win, opts) : dialog.showOpenDialog(opts))
      return r.canceled ? null : (r.filePaths[0] ?? null)
    },
  )

  ipcMain.handle(
    INTEGRATIONS_CHANNELS.saveSkillZip,
    async (_e, title: string): Promise<string | null> => {
      const skill = bundled()
      const opts: Electron.SaveDialogOptions = {
        title: String(title ?? ''),
        defaultPath: join(app.getPath('downloads'), `genoffice-skill-${skill.version}.zip`),
        filters: [{ name: 'ZIP', extensions: ['zip'] }],
      }
      const win = deps.window()
      const r = await (win ? dialog.showSaveDialog(win, opts) : dialog.showSaveDialog(opts))
      if (r.canceled || !r.filePath) return null
      writeFileSync(r.filePath, await buildSkillZip(skill))
      return r.filePath
    },
  )

  ipcMain.handle(INTEGRATIONS_CHANNELS.copyText, (_e, text: string): void => {
    if (typeof text === 'string') clipboard.writeText(text)
  })
}

function cliVersion(packageJson: string): string {
  try {
    return String(JSON.parse(readFileSync(packageJson, 'utf-8')).version ?? '')
  } catch {
    return ''
  }
}

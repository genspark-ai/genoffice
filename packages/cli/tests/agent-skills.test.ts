import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import {
  agentConfigDir,
  agentTarget,
  buildSkillZip,
  bundledSkillFrom,
  compareVersions,
  currentSkillInstalled,
  detectAgents,
  findLegacySkillCopies,
  installSkill,
  parseSkillFrontmatter,
  readInstallState,
  sweepLegacySkillCopies,
  uninstallSkill,
  type SkillLedger,
} from '../src/agent-skills'

const skillText = (version: string, body = 'Run `genoffice --version` first.') =>
  `---\nname: genoffice\ndescription: test\nmetadata:\n  version: ${version}\n  cli: '>=0.4.0'\n---\n\n${body}\n`

const bundled = (version = '1.2.0') => bundledSkillFrom(Buffer.from(skillText(version)))

function home(): string {
  return mkdtempSync(join(tmpdir(), 'genoffice-skills-'))
}

describe('detectAgents', () => {
  it('lists the tools whose dotfolder exists, honouring the env overrides', () => {
    const h = home()
    mkdirSync(join(h, '.claude'))
    mkdirSync(join(h, 'codex-elsewhere'))
    const found = detectAgents({ CODEX_HOME: join(h, 'codex-elsewhere') }, h)
    expect(found.map((a) => a.id)).toEqual(['claude-code', 'agents'])
    expect(found[0]!.skillsDir).toBe(join(h, '.claude', 'skills'))
    expect(found[1]!.skillsDir).toBe(join(h, '.agents', 'skills'))
    expect(found[1]!.label).toBe(
      '.agents (Codex / Cursor / Gemini CLI / Copilot / OpenCode / dsh …)',
    )
    expect(agentTarget('windsurf', {}, h)!.skillsDir).toBe(
      join(h, '.codeium', 'windsurf', 'skills'),
    )
  })

  it('sends the ~/.agents readers to the shared row, and keeps a probe for their dotfolders', () => {
    const h = home()
    mkdirSync(join(h, '.codex'))
    mkdirSync(join(h, '.cursor'))
    mkdirSync(join(h, '.gemini'))
    mkdirSync(join(h, '.copilot'))
    mkdirSync(join(h, '.config', 'opencode'), { recursive: true })
    expect(detectAgents({}, h).map((a) => a.id)).toEqual(['agents'])
    expect(agentTarget('codex', {}, h)).toBeNull()
    expect(agentTarget('cursor', {}, h)).toBeNull()
    expect(agentTarget('gemini', {}, h)).toBeNull()
    expect(agentTarget('copilot', {}, h)).toBeNull()
    expect(agentTarget('opencode', {}, h)).toBeNull()
    // `genoffice mcp` still reaches the per-tool dotfolders
    expect(agentConfigDir('codex', {}, h)).toBe(join(h, '.codex'))
    expect(agentConfigDir('cursor', {}, h)).toBe(join(h, '.cursor'))
    expect(agentConfigDir('gemini', {}, h)).toBe(join(h, '.gemini'))
    expect(agentConfigDir('copilot', {}, h)).toBe(join(h, '.copilot'))
    expect(agentConfigDir('opencode', {}, h)).toBe(join(h, '.config', 'opencode'))
  })

  it('shows the shared row for ~/.agents itself and keeps Windsurf and OpenClaw rows of their own', () => {
    const h = home()
    mkdirSync(join(h, '.agents', 'skills'), { recursive: true })
    mkdirSync(join(h, '.codeium', 'windsurf'), { recursive: true })
    mkdirSync(join(h, '.openclaw'))
    const found = detectAgents({}, h)
    expect(found.map((a) => a.id)).toEqual(['agents', 'windsurf', 'openclaw'])
    expect(found[0]!.skillsDir).toBe(join(h, '.agents', 'skills'))
    expect(found[1]!.skillsDir).toBe(join(h, '.codeium', 'windsurf', 'skills'))
    expect(found[2]!.skillsDir).toBe(join(h, '.openclaw', 'skills'))
  })

  it('keeps dsh a reader of the shared row', () => {
    const h = home()
    mkdirSync(join(h, '.dsh'))
    expect(detectAgents({}, h).map((a) => a.id)).toEqual(['agents'])
  })
})

describe('parseSkillFrontmatter / compareVersions', () => {
  it('reads name and metadata.version', () => {
    expect(parseSkillFrontmatter(skillText('2.1.0'))).toEqual({
      name: 'genoffice',
      version: '2.1.0',
    })
    expect(parseSkillFrontmatter('---\nname: other\n---\nx')).toEqual({
      name: 'other',
      version: '',
    })
    expect(parseSkillFrontmatter('no front matter')).toBeNull()
  })

  it('orders numeric semver and sorts junk lowest', () => {
    expect(compareVersions('1.2.0', '1.10.0')).toBe(-1)
    expect(compareVersions('2.0.0', '1.9.9')).toBe(1)
    expect(compareVersions('1.2.0', '1.2.0')).toBe(0)
    expect(compareVersions('', '0.0.1')).toBe(-1)
  })
})

describe('readInstallState', () => {
  it('walks the seven states of the design table', () => {
    const dir = join(home(), 'skills')
    const b = bundled('1.2.0')
    const ledger: SkillLedger = {}
    expect(readInstallState(dir, b, ledger).status).toBe('missing')

    // our install, current
    const path = installSkill(dir, b, ledger)
    expect(path).toBe(join(dir, 'genoffice', 'SKILL.md'))
    expect(readFileSync(path, 'utf-8')).toBe(b.text)
    expect(ledger[path]).toMatchObject({ version: '1.2.0', sha256: b.sha256, channel: 'app' })
    expect(readInstallState(dir, b, ledger)).toMatchObject({
      status: 'installed',
      installedVersion: '1.2.0',
    })

    // a newer bundle arrives
    expect(readInstallState(dir, bundled('1.3.0'), ledger)).toMatchObject({
      status: 'outdated',
      installedVersion: '1.2.0',
      older: true,
    })

    // the user edits the file
    writeFileSync(path, skillText('1.2.0', 'edited by hand'))
    expect(readInstallState(dir, b, ledger).status).toBe('modified')

    // written by someone else (no ledger entry): older → offers an update, current → nothing
    writeFileSync(path, skillText('1.1.0'))
    expect(readInstallState(dir, b, {})).toMatchObject({
      status: 'foreign',
      installedVersion: '1.1.0',
      older: true,
    })
    writeFileSync(path, skillText('1.2.0'))
    expect(readInstallState(dir, b, {})).toMatchObject({ status: 'foreign', older: false })

    // newer than the bundle, whoever wrote it
    writeFileSync(path, skillText('1.3.0'))
    expect(readInstallState(dir, b, ledger)).toMatchObject({
      status: 'newer',
      installedVersion: '1.3.0',
    })

    // the folder holds something else
    writeFileSync(path, '---\nname: other-skill\n---\n')
    expect(readInstallState(dir, b, ledger).status).toBe('occupied')
    mkdirSync(join(dir, 'genoffice', 'nested'))
    writeFileSync(path, 'not a skill at all')
    expect(readInstallState(dir, b, ledger).status).toBe('occupied')
  })
})

describe('uninstallSkill', () => {
  it('removes only what this app wrote and keeps a folder that holds other files', () => {
    const dir = join(home(), 'skills')
    const b = bundled()
    const ledger: SkillLedger = {}
    // foreign file: never touched
    mkdirSync(join(dir, 'genoffice'), { recursive: true })
    writeFileSync(join(dir, 'genoffice', 'SKILL.md'), skillText('1.0.0'))
    expect(uninstallSkill(dir, ledger)).toBe(false)
    expect(existsSync(join(dir, 'genoffice', 'SKILL.md'))).toBe(true)

    const path = installSkill(dir, b, ledger)
    writeFileSync(join(dir, 'genoffice', 'notes.txt'), 'mine')
    expect(uninstallSkill(dir, ledger)).toBe(true)
    expect(existsSync(path)).toBe(false)
    expect(existsSync(join(dir, 'genoffice', 'notes.txt'))).toBe(true)
    expect(ledger[path]).toBeUndefined()

    installSkill(dir, b, ledger)
    // an empty folder goes with the file
    const dir2 = join(home(), 'skills')
    const p2 = installSkill(dir2, b, ledger)
    expect(uninstallSkill(dir2, ledger)).toBe(true)
    expect(existsSync(join(dir2, 'genoffice'))).toBe(false)
    expect(existsSync(p2)).toBe(false)
  })
})

describe('buildSkillZip', () => {
  it('holds genoffice/SKILL.md with the exact bundled bytes', async () => {
    const b = bundled('1.2.0')
    const zip = await JSZip.loadAsync(await buildSkillZip(b))
    expect(Object.keys(zip.files).filter((f) => !zip.files[f]!.dir)).toEqual(['genoffice/SKILL.md'])
    expect(await zip.file('genoffice/SKILL.md')!.async('string')).toBe(b.text)
  })
})

describe('the legacy per-tool skill copies', () => {
  it('are found through the ledger, not by guessing', () => {
    const h = home()
    const ledger: SkillLedger = {}
    const b = bundled()
    // a copy this app wrote in the old layout
    installSkill(join(h, '.codex', 'skills'), b, ledger)
    // a copy in a folder the ledger does not know: not ours to touch
    mkdirSync(join(h, '.gemini', 'skills', 'genoffice'), { recursive: true })
    writeFileSync(join(h, '.gemini', 'skills', 'genoffice', 'SKILL.md'), 'hand written\n')

    const found = findLegacySkillCopies(ledger, {}, h)
    expect(found.map((c) => c.agent)).toEqual(['codex'])
    expect(found[0]).toMatchObject({
      label: 'Codex',
      skillsDir: join(h, '.codex', 'skills'),
      ours: true,
    })
  })

  it('sweep away the copies this app wrote and keep the edited ones', () => {
    const h = home()
    const ledger: SkillLedger = {}
    const b = bundled()
    const codexPath = installSkill(join(h, '.codex', 'skills'), b, ledger)
    const cursorPath = installSkill(join(h, '.cursor', 'skills'), b, ledger)
    // someone edited this one by hand
    writeFileSync(cursorPath, 'my own version\n')

    const sweep = sweepLegacySkillCopies(ledger, {}, h)
    expect(sweep.removed.map((c) => c.agent)).toEqual(['codex'])
    expect(sweep.kept.map((c) => c.agent)).toEqual(['cursor'])
    expect(existsSync(codexPath)).toBe(false)
    expect(existsSync(join(h, '.codex', 'skills', 'genoffice'))).toBe(false)
    expect(existsSync(cursorPath)).toBe(true)
    expect(ledger[codexPath]).toBeUndefined()
    expect(ledger[cursorPath]).toBeDefined()
  })

  it('follow the env overrides the old layout used', () => {
    const h = home()
    const ledger: SkillLedger = {}
    const b = bundled()
    const codexPath = installSkill(join(h, 'codex-elsewhere', 'skills'), b, ledger)
    const opencodePath = installSkill(join(h, 'cfg', 'opencode', 'skills'), b, ledger)

    const found = findLegacySkillCopies(
      ledger,
      { CODEX_HOME: join(h, 'codex-elsewhere'), XDG_CONFIG_HOME: join(h, 'cfg') },
      h,
    )
    expect(found.map((c) => c.path).sort()).toEqual([codexPath, opencodePath].sort())
  })

  it('the shared .agents copy gates the sweep', () => {
    const h = home()
    const ledger: SkillLedger = {}
    const b = bundled()
    installSkill(join(h, '.codex', 'skills'), b, ledger)
    expect(currentSkillInstalled(ledger, {}, h)).toBe(false)

    const sharedPath = installSkill(join(h, '.agents', 'skills'), b, ledger)
    expect(currentSkillInstalled(ledger, {}, h)).toBe(true)
    // a ledger row for a file that is gone must not open the gate
    rmSync(sharedPath)
    expect(currentSkillInstalled(ledger, {}, h)).toBe(false)
  })

  it('a ledger row whose legacy file is already gone is dropped, not kept as edited', () => {
    const h = home()
    const ledger: SkillLedger = {}
    const codexPath = installSkill(join(h, '.codex', 'skills'), bundled(), ledger)
    rmSync(codexPath)
    expect(findLegacySkillCopies(ledger, {}, h)).toEqual([])
    expect(ledger[codexPath]).toBeUndefined()
  })
})

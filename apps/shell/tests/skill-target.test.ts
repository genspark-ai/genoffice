import { describe, expect, it } from 'vitest'
import { isInstallableSkillDir } from '../src/main/skill-target'

describe('isInstallableSkillDir', () => {
  const vouched = ['/home/me/.claude/skills', '/home/me/.codex/skills', '/Volumes/work/skills']

  it('accepts a directory the user picked in the folder dialog', () => {
    expect(isInstallableSkillDir('/home/me/.claude/skills', vouched)).toBe(true)
    expect(isInstallableSkillDir('/Volumes/work/skills', vouched)).toBe(true)
  })

  it('accepts a vouched directory spelled differently', () => {
    expect(isInstallableSkillDir('/home/me/.codex/skills/../skills', vouched)).toBe(true)
    expect(isInstallableSkillDir('/home/me/.codex//skills', vouched)).toBe(true)
  })

  it('rejects a directory the main process never vouched for', () => {
    expect(isInstallableSkillDir('/etc/cron.d', vouched)).toBe(false)
    expect(isInstallableSkillDir('/Users/me/.ssh', vouched)).toBe(false)
    expect(isInstallableSkillDir('/home/me/.claude/skills-extra', vouched)).toBe(false)
    expect(isInstallableSkillDir('/home/me/.claude/skills/genoffice', vouched)).toBe(false)
  })

  it('tolerates malformed input defensively', () => {
    expect(isInstallableSkillDir('', vouched)).toBe(false)
    expect(isInstallableSkillDir('relative/dir', vouched)).toBe(false)
    expect(isInstallableSkillDir(undefined, vouched)).toBe(false)
    expect(isInstallableSkillDir('/etc/cron.d', [])).toBe(false)
  })
})

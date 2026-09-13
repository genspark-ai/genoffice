import { execFileSync, spawnSync } from 'node:child_process'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { beforeEach, describe, expect, it } from 'vitest'
import { tempDir } from './helpers'

const SCRIPT = join(__dirname, '../../../tools/check-skill-version.mjs')

let repo: string

function git(...args: string[]): void {
  execFileSync('git', args, { cwd: repo, stdio: 'ignore' })
}

function skill(name: string, version: string, body: string): string {
  return `---\nname: ${name}\ndescription: test\nmetadata:\n  version: ${version}\n---\n\n${body}\n`
}

function writeSkill(dir: string, content: string): void {
  mkdirSync(join(repo, 'skills', dir), { recursive: true })
  writeFileSync(join(repo, 'skills', dir, 'SKILL.md'), content)
}

function check(): { ok: boolean; stderr: string } {
  const r = spawnSync(process.execPath, [SCRIPT, '--base', 'HEAD'], {
    cwd: repo,
    encoding: 'utf-8',
  })
  return { ok: r.status === 0, stderr: r.stderr }
}

describe('check-skill-version', () => {
  beforeEach(() => {
    repo = tempDir()
    git('init', '-q')
    git('config', 'user.email', 'ci@example.com')
    git('config', 'user.name', 'ci')
    writeSkill('demo', skill('demo', '1.0.0', 'Run `demo --help`.'))
    git('add', '.')
    git('commit', '-q', '-m', 'base')
  })

  it('passes when nothing changed and when the body change comes with a bump', () => {
    expect(check().ok).toBe(true)
    writeSkill('demo', skill('demo', '1.1.0', 'Run `demo --help` first.'))
    expect(check().ok).toBe(true)
  })

  it('fails a body change that keeps the version', () => {
    writeSkill('demo', skill('demo', '1.0.0', 'Run `demo --help` first.'))
    const r = check()
    expect(r.ok).toBe(false)
    expect(r.stderr).toContain('bump it')
  })

  it('ignores frontmatter-only edits and rejects a version going backwards', () => {
    writeSkill('demo', skill('demo', '1.0.0', 'Run `demo --help`.').replace('test', 'demo tool'))
    expect(check().ok).toBe(true)
    writeSkill('demo', skill('demo', '0.9.0', 'Run `demo --help`.'))
    expect(check().stderr).toContain('went backwards')
  })

  it('requires a version and a name matching the directory on new skills', () => {
    writeSkill('other', '---\nname: other\ndescription: x\n---\n\nbody\n')
    expect(check().stderr).toContain('needs metadata.version')
    writeSkill('other', skill('mismatch', '1.0.0', 'body'))
    expect(check().stderr).toContain('must match its directory')
  })
})

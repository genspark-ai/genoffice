import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { genofficeUserDataDir, guiOpenDocuments } from '../src/gui'
import { run, tempDir } from './helpers'

function writeRegistry(dir: string, pid: number, paths: string[]): void {
  mkdirSync(dir, { recursive: true })
  writeFileSync(
    join(dir, 'open-documents.json'),
    JSON.stringify({ pid, updatedAt: new Date().toISOString(), paths }),
  )
}

function registry(dir: string, pid: number, paths: string[]): Record<string, string> {
  writeRegistry(dir, pid, paths)
  return { ...process.env, GENOFFICE_AUDIT_LOG: 'off', GENOFFICE_USER_DATA: dir }
}

async function workbook(dir: string): Promise<string> {
  const table = join(dir, 't.json')
  writeFileSync(table, JSON.stringify([['a'], [1]]))
  const xlsx = join(dir, 't.xlsx')
  const r = await run(['create', '--type', 'xlsx', '--from', table, '--out', xlsx], {
    env: { ...process.env, GENOFFICE_AUDIT_LOG: 'off' },
  })
  expect(r.code).toBe(0)
  return xlsx
}

describe('GUI-open documents', () => {
  it('locates the shell userData directory and honours the override', () => {
    expect(genofficeUserDataDir({ GENOFFICE_USER_DATA: '/u' })).toBe('/u')
    expect(genofficeUserDataDir({}).endsWith('GenOffice')).toBe(true)
  })

  it('ignores a missing, malformed or crash-leftover registry', () => {
    const dir = tempDir()
    expect(guiOpenDocuments({ GENOFFICE_USER_DATA: dir })).toEqual([])
    writeFileSync(join(dir, 'open-documents.json'), '{not json')
    expect(guiOpenDocuments({ GENOFFICE_USER_DATA: dir })).toEqual([])
    registry(dir, 2 ** 22 + 12345, ['/x.docx'])
    expect(guiOpenDocuments({ GENOFFICE_USER_DATA: dir })).toEqual([])
    registry(dir, process.pid, ['/x.docx'])
    expect(guiOpenDocuments({ GENOFFICE_USER_DATA: dir })).toEqual([
      { pid: process.pid, paths: ['/x.docx'] },
    ])
  })

  it('refuses a file open in the dev registry after finding a live packaged registry', async () => {
    const dir = tempDir()
    const xlsx = await workbook(dir)
    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'set_range', range: 'A1', values: [['z']] }]))
    const appRoot = tempDir()
    const env: NodeJS.ProcessEnv = {
      ...process.env,
      APPDATA: appRoot,
      GENOFFICE_AUDIT_LOG: 'off',
      HOME: appRoot,
      XDG_CONFIG_HOME: appRoot,
    }
    delete env.GENOFFICE_USER_DATA
    const packagedDir = genofficeUserDataDir(env)
    writeRegistry(packagedDir, process.pid, [join(dir, 'packaged.xlsx')])
    writeRegistry(`${packagedDir} Dev`, process.pid, [xlsx])

    const refused = await run(['sheet', 'apply', xlsx, '--ops', ops, '--json'], { env })

    expect(refused.code).toBe(2)
    expect(refused.json().detail).toMatchObject({ gui_pid: process.pid })
  })

  it('refuses an in-place edit of an open file unless --force, other files pass', async () => {
    const dir = tempDir()
    const xlsx = await workbook(dir)
    const ops = join(dir, 'ops.json')
    writeFileSync(ops, JSON.stringify([{ op: 'set_range', range: 'A1', values: [['z']] }]))
    const env = registry(join(dir, 'userData'), process.pid, [xlsx])
    const refused = await run(['sheet', 'apply', xlsx, '--ops', ops, '--json'], { env })
    expect(refused.code).toBe(2)
    expect(refused.json().message).toContain('has this file open')
    expect(refused.json().detail.gui_pid).toBe(process.pid)
    const elsewhere = await run(
      ['sheet', 'apply', xlsx, '--ops', ops, '--out', join(dir, 'copy.xlsx'), '--json'],
      { env },
    )
    expect(elsewhere.code).toBe(0)
    const forced = await run(['sheet', 'apply', xlsx, '--ops', ops, '--force', '--json'], { env })
    expect(forced.code).toBe(0)
  })

  it('applies to convert and create outputs too', async () => {
    const dir = tempDir()
    const csv = join(dir, 'a.csv')
    writeFileSync(csv, 'x\n1\n')
    const env = registry(join(dir, 'userData'), process.pid, [join(dir, 'a.xlsx')])
    expect((await run(['convert', csv, '--to', 'xlsx', '--json'], { env })).code).toBe(2)
    expect((await run(['convert', csv, '--to', 'xlsx', '--force', '--json'], { env })).code).toBe(0)
  })
})

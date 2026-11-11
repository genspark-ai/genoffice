import { describe, expect, it, vi } from 'vitest'
import {
  OFFICE_TYPES,
  createDefaultAppService,
  decodeCommandOutput,
  macAppBundlePath,
  parseMacStatus,
  parseRegValue,
} from '../src/main/default-app'
import type { RunCommand } from '../src/main/default-app'

const ME = 'com.genoffice.app'
const WPS = 'com.kingsoft.wpsoffice.mac'

function macStatusJson(
  owner: (uti: string) => string | null,
  name: string | null = 'WPS Office',
): string {
  return JSON.stringify({
    me: ME,
    types: OFFICE_TYPES.map((t) => {
      const id = owner(t.uti)
      return { uti: t.uti, id, name: id === ME ? 'GenOffice' : id ? name : null }
    }),
  })
}

describe('macAppBundlePath', () => {
  it('walks up from the executable to the .app bundle', () => {
    expect(macAppBundlePath('/Applications/GenOffice.app/Contents/MacOS/GenOffice')).toBe(
      '/Applications/GenOffice.app',
    )
    expect(macAppBundlePath('/usr/local/bin/electron')).toBeNull()
  })
})

describe('parseMacStatus', () => {
  it('is default when every type resolves to our bundle id', () => {
    expect(parseMacStatus(macStatusJson(() => ME))).toEqual({
      state: 'default',
      others: [],
      manualOnly: false,
    })
  })

  it('names the other app once even when it owns several types', () => {
    expect(parseMacStatus(macStatusJson(() => WPS))).toEqual({
      state: 'other',
      others: ['WPS Office'],
      manualOnly: false,
    })
  })

  it('is "other" when a single type belongs to someone else', () => {
    const json = macStatusJson((uti) => (uti.includes('presentationml') ? WPS : ME))
    expect(parseMacStatus(json).state).toBe('other')
  })

  it('falls back to the bundle id when the owner has no display name', () => {
    expect(parseMacStatus(macStatusJson(() => 'com.example.x', null)).others).toEqual([
      'com.example.x',
    ])
  })

  it('treats a type with no handler as not yet claimed', () => {
    const json = macStatusJson((uti) => (uti.includes('macroenabled') ? null : ME))
    expect(parseMacStatus(json)).toEqual({ state: 'other', others: [], manualOnly: false })
  })

  it('is unknown when LaunchServices has no handler at all', () => {
    expect(parseMacStatus(macStatusJson(() => null)).state).toBe('unknown')
  })
})

describe('parseRegValue', () => {
  it('reads the data column of a reg query line', () => {
    const out = [
      '',
      'HKEY_CURRENT_USER\\Software\\...\\.docx\\UserChoice',
      '    ProgId    REG_SZ    Word.Document.12',
      '',
    ].join('\r\n')
    expect(parseRegValue(out)).toBe('Word.Document.12')
    expect(parseRegValue('    (Default)    REG_SZ    Word Document\r\n')).toBe('Word Document')
    expect(parseRegValue('    (Default)    REG_SZ    Microsoft Excel 工作表\r\n')).toBe(
      'Microsoft Excel 工作表',
    )
    expect(
      parseRegValue('ERROR: The system was unable to find the specified registry key'),
    ).toBeNull()
  })
})

describe('decodeCommandOutput', () => {
  // A Chinese display name in GBK, the code page reg.exe uses on a Chinese Windows install.
  const gbkName = Buffer.concat([
    Buffer.from('Microsoft Excel ', 'ascii'),
    Buffer.from([0xb9, 0xa4, 0xd7, 0xf7, 0xb1, 0xed]),
  ])

  it('recovers the localised ProgId name instead of mojibake on Windows', () => {
    expect(gbkName.toString('utf8')).toContain('\uFFFD') // the old behaviour
    expect(decodeCommandOutput(gbkName, 'win32')).toBe('Microsoft Excel 工作表')
  })

  it('leaves ASCII output untouched on Windows', () => {
    const ascii = Buffer.from('    (Default)    REG_SZ    Word.Document.12\r\n', 'ascii')
    expect(decodeCommandOutput(ascii, 'win32')).toBe(ascii.toString('utf8'))
  })

  it('keeps genuine UTF-8 output on Windows', () => {
    const utf8 = Buffer.from('Microsoft PowerPoint 演示文稿', 'utf8')
    expect(decodeCommandOutput(utf8, 'win32')).toBe('Microsoft PowerPoint 演示文稿')
  })

  it('uses UTF-8 on macOS and Linux', () => {
    const utf8 = Buffer.from('WPS Office — 表格', 'utf8')
    expect(decodeCommandOutput(utf8, 'linux')).toBe('WPS Office — 表格')
    expect(decodeCommandOutput(utf8, 'darwin')).toBe('WPS Office — 表格')
  })

  it('passes through an already-decoded string', () => {
    expect(decodeCommandOutput('Word Document', 'win32')).toBe('Word Document')
  })
})

describe('createDefaultAppService', () => {
  const base = {
    packaged: true,
    exePath: '/Applications/GenOffice.app/Contents/MacOS/GenOffice',
    openExternal: vi.fn(async () => {}),
  }

  it('probes read-only in dev builds but never claims the types', async () => {
    const calls: string[][] = []
    const run = vi.fn<RunCommand>(async (cmd, args) => {
      calls.push([cmd, ...args])
      return macStatusJson(() => WPS)
    })
    const svc = createDefaultAppService({ ...base, platform: 'darwin', packaged: false, run })
    expect((await svc.status()).state).toBe('other')
    // set() re-probes instead of writing, so the settings entry stays visible.
    expect((await svc.set()).state).toBe('other')
    expect(calls).toHaveLength(2)
    expect(
      calls.some((c) => c.some((a) => a.includes('LSSetDefaultRoleHandlerForContentType'))),
    ).toBe(false)
  })

  it('mac: claims every UTI through osascript and re-reads the status', async () => {
    let owner = WPS
    const run = vi.fn<RunCommand>(async (cmd, args) => {
      expect(cmd).toBe('osascript')
      expect(args.slice(0, 3)).toEqual(['-l', 'JavaScript', '-e'])
      expect(args[4]).toBe('/Applications/GenOffice.app')
      expect(args.slice(5)).toEqual(OFFICE_TYPES.map((t) => t.uti))
      if (args[3].includes('LSSetDefaultRoleHandlerForContentType')) {
        owner = ME
        return JSON.stringify({ me: ME, failed: [] })
      }
      return macStatusJson(() => owner)
    })
    const svc = createDefaultAppService({ ...base, platform: 'darwin', run })
    expect(await svc.status()).toMatchObject({ state: 'other', others: ['WPS Office'] })
    expect(await svc.set()).toMatchObject({ state: 'default', others: [] })
    expect(run).toHaveBeenCalledTimes(3)
  })

  it('linux: compares xdg-mime answers with our desktop id', async () => {
    const calls: string[][] = []
    const run = vi.fn<RunCommand>(async (cmd, args) => {
      calls.push([cmd, ...args])
      if (args[0] === 'query')
        return args[2].includes('spreadsheetml') ? 'wps-office-et.desktop\n' : 'genoffice.desktop\n'
      return ''
    })
    const svc = createDefaultAppService({
      ...base,
      platform: 'linux',
      run,
      readFile: (p) => {
        // path.join uses backslashes on Windows, so match the file name only.
        if (p.includes('wps-office-et.desktop')) return '[Desktop Entry]\nName=WPS Spreadsheets\n'
        throw new Error('ENOENT')
      },
    })
    expect(await svc.status()).toEqual({
      state: 'other',
      others: ['WPS Spreadsheets'],
      manualOnly: false,
    })
    await svc.set()
    expect(calls.find((c) => c[1] === 'default')).toEqual([
      'xdg-mime',
      'default',
      'genoffice.desktop',
      ...OFFICE_TYPES.map((t) => t.mime),
    ])
  })

  it('windows: reads UserChoice ProgIds, names them, and only opens the settings page', async () => {
    const run = vi.fn<RunCommand>(async (_cmd, args) => {
      const key = args[1]
      if (key.includes('UserChoice')) {
        if (key.includes('.pptx')) throw new Error('no user choice')
        return key.includes('.docx')
          ? '    ProgId    REG_SZ    Word.Document.12\r\n'
          : '    ProgId    REG_SZ    Excel Workbook\r\n'
      }
      if (key === 'HKCR\\.pptx') return '    (Default)    REG_SZ    PowerPoint Presentation\r\n'
      if (key === 'HKCR\\Word.Document.12')
        return '    (Default)    REG_SZ    Microsoft Word Document\r\n'
      throw new Error('missing')
    })
    const openExternal = vi.fn(async () => {})
    const svc = createDefaultAppService({ ...base, platform: 'win32', run, openExternal })
    expect(await svc.status()).toEqual({
      state: 'other',
      others: ['Microsoft Word Document'],
      manualOnly: true,
    })
    await svc.set()
    expect(openExternal).toHaveBeenCalledWith('ms-settings:defaultapps')
  })

  it('windows: falls back to the ProgId when a name still has replacement characters', async () => {
    const run = vi.fn<RunCommand>(async (_cmd, args) => {
      const key = args[1]
      if (key.includes('UserChoice')) return '    ProgId    REG_SZ    Excel.Sheet.12\r\n'
      if (key === 'HKCR\\Excel.Sheet.12')
        return '    (Default)    REG_SZ    Microsoft Excel \uFFFD\uFFFD\uFFFD\r\n'
      throw new Error('missing')
    })
    const svc = createDefaultAppService({
      ...base,
      platform: 'win32',
      run,
      openExternal: vi.fn(async () => {}),
    })
    expect((await svc.status()).others).toEqual(['Excel.Sheet.12'])
  })

  it('reports unknown instead of throwing when the probe fails', async () => {
    const run = vi.fn<RunCommand>(async () => {
      throw new Error('osascript missing')
    })
    const svc = createDefaultAppService({ ...base, platform: 'darwin', run })
    expect((await svc.status()).state).toBe('unknown')
  })
})

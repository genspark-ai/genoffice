import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { appBinaryForResources } from '../src/resources'

describe('appBinaryForResources', () => {
  it('finds the app in a custom Windows install directory', () => {
    expect(appBinaryForResources(join('D:\\Apps\\GenOffice', 'resources'), 'win32')).toBe(
      join('D:\\Apps\\GenOffice', 'GenOffice.exe'),
    )
  })

  it('resolves the same binary for the default Windows install directory', () => {
    const localAppData = 'C:\\Users\\test\\AppData\\Local'
    const resources = join(localAppData, 'Programs', 'GenOffice', 'resources')
    expect(appBinaryForResources(resources, 'win32')).toBe(
      join(localAppData, 'Programs', 'GenOffice', 'GenOffice.exe'),
    )
  })

  it('finds the app in a custom macOS bundle location', () => {
    expect(
      appBinaryForResources(join('/Volumes/Work/GenOffice.app/Contents/Resources'), 'darwin'),
    ).toBe(join('/Volumes/Work/GenOffice.app/Contents/MacOS/GenOffice'))
  })

  it('finds the app in a custom Linux prefix', () => {
    expect(appBinaryForResources(join('/opt/genoffice-custom/resources'), 'linux')).toBe(
      join('/opt/genoffice-custom/genoffice'),
    )
  })
})

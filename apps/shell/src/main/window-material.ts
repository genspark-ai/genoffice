import { release } from 'node:os'

/// Windows 11 22H2 (build 22621) is the first release where Electron's
/// `backgroundMaterial` takes effect; earlier builds ignore it and would paint
/// an opaque black window behind a transparent page.
const WIN11_22H2_BUILD = 22621

/// True when the shell window gets the Mica material behind its transparent page.
export function winMaterialSupported(
  platform: NodeJS.Platform = process.platform,
  osRelease: string = release(),
): boolean {
  if (platform !== 'win32') return false
  const build = Number(osRelease.split('.')[2])
  return Number.isFinite(build) && build >= WIN11_22H2_BUILD
}

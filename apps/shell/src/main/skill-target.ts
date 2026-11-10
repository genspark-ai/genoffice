import { isAbsolute, resolve } from 'node:path'

export function isInstallableSkillDir(dir: unknown, vouched: Iterable<string>): dir is string {
  if (typeof dir !== 'string' || dir === '' || !isAbsolute(dir)) return false
  const key = resolve(dir)
  for (const allowed of vouched) if (resolve(allowed) === key) return true
  return false
}

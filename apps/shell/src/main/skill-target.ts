/**
 * Gate for the skill-install IPC (apps/shell/src/main/integrations-ipc.ts).
 *
 * installSkill writes `<dir>/genoffice/SKILL.md`, and the only directory the
 * renderer ever names is the one it got back from the folder dialog. A
 * compromised renderer can pass any path instead, so the write is constrained
 * to the directories the main process vouched for: the agents' own skills
 * directories, plus the folder the user just picked. No single root can express
 * that — an agent dir follows CLAUDE_CONFIG_DIR / XDG_CONFIG_HOME and may sit
 * outside the home directory, while a picked folder is wherever the user keeps
 * their dotfolders.
 */

import { isAbsolute, resolve } from 'node:path'

/** `dir` must name one of `vouched` exactly, compared resolved, or nothing is written. */
export function isInstallableSkillDir(dir: unknown, vouched: Iterable<string>): dir is string {
  if (typeof dir !== 'string' || dir === '' || !isAbsolute(dir)) return false
  const key = resolve(dir)
  for (const allowed of vouched) if (resolve(allowed) === key) return true
  return false
}

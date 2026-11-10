/**
 * The shape of a SKILL.md GenOffice found on this machine, shared by the main
 * process that reads the files and the renderer that lists them.
 *
 * It lives here rather than beside the reader (`main/imported-skills.ts`) so the
 * renderer never has to name a main-process module: the reader reaches for
 * `node:fs`, and a `import type` across that boundary is erased today but reads
 * like a bundling mistake tomorrow.
 */

import type { SkillRelevance } from '@genoffice/agent-core'

/** one SKILL.md we found, with everything the palette needs to show it */
export interface FoundSkill {
  /** the directory name, which is what `skills/<name>/` is keyed on */
  name: string
  description: string
  /** absolute path of the SKILL.md, read when the user picks it */
  path: string
  /** whether it talks about a format GenOffice opens, and on what */
  relevance: SkillRelevance
}

/** a skill GenOffice ships as a suggestion, with where the user stands on it */
export interface RecommendedSkill extends FoundSkill {
  version: string
  /** the frontmatter's `category`; groups the Recommended list (office, finance, ...) */
  category: string
  /** missing: not installed; installed: same version in our folder; outdated: ours is older */
  state: 'missing' | 'installed' | 'outdated'
}

/**
 * The skills a user attached to one message from the composer.
 *
 * Skills are never offered to the model on their own: the only way a skill's
 * text reaches a turn is the user picking it, and then the whole body rides in
 * front of the request. The display text stays the user's own words; this is
 * the model-facing instruction.
 */
export interface AttachedSkill {
  name: string
  body: string
}

export function withAttachedSkills(instruction: string, skills: readonly AttachedSkill[]): string {
  if (skills.length === 0) return instruction
  const blocks = skills
    .map((s) => `<skill name="${s.name.replace(/"/g, '')}">\n${s.body.trim()}\n</skill>`)
    .join('\n\n')
  return (
    `The user attached the following skill${skills.length > 1 ? 's' : ''} to this request. ` +
    `Follow their steps and constraints with the tools you have; where a skill and general ` +
    `guidance disagree, the skill wins.\n\n${blocks}\n\nUser request:\n${instruction}`
  )
}

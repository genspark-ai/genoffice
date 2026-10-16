import type { FeatureRecord } from './features-skill'
import { validateScript, type FeatureScript } from './feature-script'

/**
 * The hook kind's contract (K2): a feature that observes an event the app raises.
 *
 * The events are the ones the host itself can see and name — a document opening or
 * closing. A hook runs in the same sandbox as a script, it is handed the event as
 * data, and it cannot write: whatever it wants to say goes to its log and to the
 * feature panel's activity list. One event — document.will-save — is the host asking
 * permission, so a hook may also call $.event.veto(reason); the host then refuses
 * that save and shows the reason. Nothing else a hook does can change anything.
 */

export const HOOK_EVENTS = ['document.opened', 'document.closed', 'document.will-save'] as const
export type HookEvent = (typeof HOOK_EVENTS)[number]

export interface FeatureHook {
  /** the event this hook observes */
  on: HookEvent
  /** the same body a script uses; it receives ($, input) with input.event set */
  script: FeatureScript
}

export interface HookValidationError {
  field: string
  message: string
}

export type HookValidation =
  { ok: true; hook: FeatureHook } | { ok: false; errors: HookValidationError[] }

/** Validate a hook payload from the assistant or a mod file. */
export function validateHook(value: unknown): HookValidation {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return { ok: false, errors: [{ field: 'hook', message: 'needs a hook object' }] }
  }
  const input = value as Record<string, unknown>
  const errors: HookValidationError[] = []
  if (typeof input.on !== 'string' || !(HOOK_EVENTS as readonly string[]).includes(input.on)) {
    errors.push({
      field: 'hook.on',
      message: 'must be an event the host raises (' + HOOK_EVENTS.join(', ') + ')',
    })
  }
  const script = validateScript(input.script)
  if (!script.ok) {
    for (const error of script.errors) {
      errors.push({ field: 'hook.' + error.field, message: error.message })
    }
  }
  if (errors.length) return { ok: false, errors }
  return {
    ok: true,
    hook: { on: input.on as HookEvent, script: script.ok ? script.script : { source: '' } },
  }
}

/** The hooks that observe one event, in registry order. */
export function hooksFor(features: readonly FeatureRecord[], event: HookEvent): FeatureRecord[] {
  return features.filter((feature) => feature.hook?.on === event)
}

/** One line per hook, for the assistant's context. */
export function describeHooks(features: readonly FeatureRecord[]): string[] {
  return features
    .filter((feature) => feature.hook)
    .map((feature) => feature.id + ' observes ' + feature.hook!.on)
}

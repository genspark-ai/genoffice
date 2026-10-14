import type { FeatureRecord } from './features-skill'
import { validateSlotAction, type SlotAction, type SlotValidationError } from './feature-slots'

/**
 * The command kind's contract (K6): an action the user can invoke from a list.
 *
 * A command is deliberately the *same* declared actions a drawn control may run —
 * open a file, hand the assistant a prompt, open a feature's own panel — invoked
 * from a palette instead of from a button. Nothing new becomes possible: the kind
 * decides where a press starts, not what it may reach.
 */

export const COMMAND_LABEL_MAX = 60

export interface FeatureCommand {
  /** what the palette lists */
  label: string
  /** one line under the label */
  description?: string
  action: SlotAction
}

export interface CommandValidationError {
  field: string
  message: string
}

export type CommandValidation =
  { ok: true; command: FeatureCommand } | { ok: false; errors: CommandValidationError[] }

/** Validate a command payload from the assistant or a mod file. */
export function validateCommand(value: unknown): CommandValidation {
  const errors: CommandValidationError[] = []
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return { ok: false, errors: [{ field: 'command', message: 'needs a command object' }] }
  }
  const input = value as Record<string, unknown>
  const unknown = Object.keys(input).filter(
    (key) => key !== 'label' && key !== 'description' && key !== 'action',
  )
  for (const key of unknown) {
    errors.push({ field: 'command.' + key, message: 'is not a command field' })
  }
  if (typeof input.label !== 'string' || input.label.trim().length === 0) {
    errors.push({ field: 'command.label', message: 'needs the text the palette lists' })
  } else if (input.label.length > COMMAND_LABEL_MAX) {
    errors.push({
      field: 'command.label',
      message: 'is longer than ' + COMMAND_LABEL_MAX + ' characters',
    })
  }
  if (input.description !== undefined) {
    if (typeof input.description !== 'string' || input.description.trim().length === 0) {
      errors.push({ field: 'command.description', message: 'needs a non-empty string' })
    } else if (input.description.length > 160) {
      errors.push({ field: 'command.description', message: 'is longer than 160 characters' })
    }
  }
  const actionErrors: SlotValidationError[] = []
  validateSlotAction(input.action, 'command.action', actionErrors)
  errors.push(...actionErrors)
  if (errors.length) return { ok: false, errors }
  return {
    ok: true,
    command: {
      label: (input.label as string).trim(),
      ...(input.description === undefined
        ? {}
        : { description: (input.description as string).trim() }),
      action: input.action as SlotAction,
    },
  }
}

/** The commands a registry holds, in registry order. */
export function commandsIn(features: readonly FeatureRecord[]): FeatureRecord[] {
  return features.filter((feature) => feature.command)
}

/**
 * Whether a command answers a query. Case-insensitive, over the label and the
 * description, and an empty query matches everything — the palette's own rule, kept
 * here so it is unit-tested rather than buried in a component.
 */
export function matchesCommand(command: FeatureCommand, query: string): boolean {
  const needle = query.trim().toLowerCase()
  if (!needle) return true
  const haystack = (command.label + ' ' + (command.description ?? '')).toLowerCase()
  return haystack.includes(needle)
}

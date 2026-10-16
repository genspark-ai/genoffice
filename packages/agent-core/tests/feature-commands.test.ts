import { describe, expect, it } from 'vitest'
import { commandsIn, matchesCommand, validateCommand } from '../src/feature-commands'
import type { FeatureRecord } from '../src/features-skill'

function record(id: string, command?: FeatureRecord['command']): FeatureRecord {
  return {
    id,
    name: id,
    description: 'x',
    kind: command ? 'command' : '',
    ...(command ? { command } : {}),
    createdAt: 'x',
    updatedAt: 'x',
  }
}

describe('command contract', () => {
  it('accepts the declared actions, and nothing else', () => {
    for (const action of [
      { kind: 'open', path: '/tmp/a.docx' },
      { kind: 'ask-assistant', prompt: 'tidy the headings' },
      { kind: 'panel', id: 'file-sidebar' },
      { kind: 'run-feature', id: 'tidy-headings' },
    ]) {
      expect(validateCommand({ label: 'Do it', action }).ok).toBe(true)
    }
    expect(validateCommand({ label: 'Do it', action: { kind: 'shell', cmd: 'rm -rf /' } }).ok).toBe(
      false,
    )
  })

  it('refuses a missing label, an over-long one, an unknown field and a bad action', () => {
    const noLabel = validateCommand({ action: { kind: 'open', path: '/tmp/a' } })
    expect(noLabel.ok).toBe(false)
    if (!noLabel.ok) expect(noLabel.errors[0]!.field).toBe('command.label')

    expect(validateCommand({ label: 'x'.repeat(61), action: { kind: 'panel', id: 'p' } }).ok).toBe(
      false,
    )
    const extra = validateCommand({
      label: 'Do it',
      action: { kind: 'open', path: '/tmp/a' },
      keys: 'Cmd+K',
    })
    expect(extra.ok).toBe(false)
    if (!extra.ok) expect(extra.errors[0]!.field).toBe('command.keys')

    const badAction = validateCommand({ label: 'Do it', action: { kind: 'open' } })
    expect(badAction.ok).toBe(false)
    if (!badAction.ok) expect(badAction.errors[0]!.field).toBe('command.action.path')
  })

  it('lists the commands a registry holds', () => {
    const features = [
      record('tidy', { label: 'Tidy headings', action: { kind: 'ask-assistant', prompt: 'tidy' } }),
      record('plain'),
    ]
    expect(commandsIn(features).map((f) => f.id)).toEqual(['tidy'])
  })

  it('matches a query over the label and the description, case-insensitively', () => {
    const command = {
      label: 'Tidy headings',
      description: 'unfold every wrapped heading',
      action: { kind: 'ask-assistant' as const, prompt: 'tidy' },
    }
    expect(matchesCommand(command, '')).toBe(true)
    expect(matchesCommand(command, '  ')).toBe(true)
    expect(matchesCommand(command, 'TIDY')).toBe(true)
    expect(matchesCommand(command, 'wrapped')).toBe(true)
    expect(matchesCommand(command, 'report')).toBe(false)
  })
})

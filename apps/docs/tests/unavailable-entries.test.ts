/**
 * An entry whose feature is not built is not shown. This holds the rule over
 * the sources: no planned command in the menu bar, the ribbon or the context
 * menu, and no control in a renderer component that is disabled whatever the
 * state.
 *
 * Limits: it reads the literal forms only (a bare `disabled` attribute,
 * `disabled={true}`, `disabled: true`). A control disabled through a helper's
 * argument or an expression that happens to be always true is not seen.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { dirname, join, relative } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { RIBBON_COMMAND_IDS } from '../src/renderer/commands/ribbon-controls'
import { COMMANDS, isPlanned, type CommandDef } from '../src/shared/commands'
import { DOCS_MENU, type MenuNode } from '../src/shared/menu-layout'
import { UNAVAILABLE_ENTRIES, isUnavailable } from '../src/shared/unavailable-entries'

const RENDERER = join(dirname(fileURLToPath(import.meta.url)), '../src/renderer')

const sources = (dir: string): string[] =>
  readdirSync(dir, { withFileTypes: true }).flatMap((entry) =>
    entry.isDirectory()
      ? entry.name === 'i18n'
        ? []
        : sources(join(dir, entry.name))
      : entry.name.endsWith('.tsx')
        ? [join(dir, entry.name)]
        : [],
  )

const FILES = sources(RENDERER).map((file) => ({
  name: relative(RENDERER, file),
  lines: readFileSync(file, 'utf8').split('\n'),
}))

const PLANNED = Object.keys(COMMANDS).filter(isPlanned)

/** a control that is disabled whatever the state, in the forms this test can read */
const ALWAYS_DISABLED = [
  /^\s*disabled\s*$/,
  /<[A-Za-z][^<>]*\sdisabled(\s|\/?>|$)/,
  /\bdisabled=\{true\}/,
  /\bdisabled: true\b/,
]
/** a read-only value or a state row, with the reason after the colon */
const EXEMPT = /always-disabled: \S/
const GUARD_OPEN = /<Unavailable id=(?:"([^"]+)"|\{)/

function alwaysDisabled(lines: readonly string[]): number[] {
  const found: number[] = []
  let guarded = false
  lines.forEach((line, i) => {
    const open = GUARD_OPEN.exec(line)
    // a guard whose row is gone no longer hides its content
    if (open) guarded = open[1] === undefined || isUnavailable(open[1])
    if (line.includes('</Unavailable>')) guarded = false
    if (guarded || !ALWAYS_DISABLED.some((form) => form.test(line))) return
    if (!lines.slice(Math.max(0, i - 3), i + 1).some((near) => EXEMPT.test(near))) found.push(i + 1)
  })
  return found
}

describe('entries whose feature is not built', () => {
  it('reads the forms of an always-disabled control', () => {
    expect(alwaysDisabled(['<button disabled>', '<input', '  disabled', '/>'])).toEqual([1, 3])
    expect(alwaysDisabled(['<input type="checkbox" disabled readOnly />'])).toEqual([1])
    expect(alwaysDisabled(['item(label, { disabled: true })', '<b disabled={true} />'])).toEqual([
      1, 2,
    ])
    expect(alwaysDisabled(['<button disabled={!hasDoc}>', 'const { disabled } = props'])).toEqual(
      [],
    )
    expect(alwaysDisabled(['// always-disabled: a computed value', '<input disabled />'])).toEqual(
      [],
    )
    const [id] = Object.keys(UNAVAILABLE_ENTRIES)
    expect(
      alwaysDisabled([`<Unavailable id="${id}">`, '<a disabled />', '</Unavailable>']),
    ).toEqual([])
    expect(
      alwaysDisabled(['<Unavailable id="no.row">', '<a disabled />', '</Unavailable>']),
    ).toEqual([2])
  })

  it('no renderer component has a control that is disabled whatever the state', () => {
    const found = FILES.flatMap((file) =>
      alwaysDisabled(file.lines).map((n) => `${file.name}:${n}`),
    )
    expect(found).toEqual([])
  })

  it('every row of the table hides something', () => {
    const text = FILES.map((file) => file.lines.join('\n')).join('\n')
    const unused = Object.keys(UNAVAILABLE_ENTRIES).filter(
      (id) => !text.includes(`"${id}"`) && !text.includes(`'${id}'`),
    )
    expect(unused).toEqual([])
  })

  it('the menu bar lists no planned command and leaves no separator behind', () => {
    const walk = (nodes: readonly MenuNode[], trail: string): void => {
      expect(nodes.length, trail).toBeGreaterThan(0)
      expect(nodes[0], trail).not.toBe('---')
      expect(nodes.at(-1), trail).not.toBe('---')
      nodes.forEach((node, i) => {
        if (node === '---') expect(nodes[i + 1], trail).not.toBe('---')
        else if (typeof node === 'string') expect(isPlanned(node), node).toBe(false)
        else if ('cmd' in node) expect(isPlanned(node.cmd), node.cmd).toBe(false)
        else if ('submenu' in node) walk(node.items, `${trail} > ${node.submenu}`)
      })
    }
    expect(PLANNED.length).toBeGreaterThan(0)
    for (const top of DOCS_MENU) walk(top.items, top.label)
    const edit = DOCS_MENU.find((top) => top.label === 'menuEdit')!.items
    const find = edit.findIndex(
      (node) => typeof node === 'object' && 'submenu' in node && node.submenu === 'menuFindMenu',
    )
    expect(edit.slice(find + 1, find + 3)).toEqual(['---', 'edit-select-data'])
  })

  it('the ribbon and the context menu offer no planned command', () => {
    expect(RIBBON_COMMAND_IDS.filter((id) => (COMMANDS[id] as CommandDef).planned)).toEqual([])
    const found: string[] = []
    for (const file of FILES) {
      if (!/(^|\/)(Ribbon|ribbon-[\w-]+|ContextMenu)\.tsx$/.test(file.name)) continue
      file.lines.forEach((line, i) => {
        for (const id of PLANNED) {
          if (!line.includes(`'${id}'`) && !line.includes(`"${id}"`)) continue
          const guard = `isPlanned('${id}')`
          if (!file.lines.slice(Math.max(0, i - 3), i + 1).some((near) => near.includes(guard)))
            found.push(`${file.name}:${i + 1} ${id}`)
        }
      })
    }
    expect(found).toEqual([])
  })
})

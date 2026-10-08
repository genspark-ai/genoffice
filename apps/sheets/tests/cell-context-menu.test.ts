import { describe, expect, it } from 'vitest'
import { ContextMenuGroup } from '@univerjs/ui'

import {
  CLEAR_FILTER_MENU_ID,
  FILTER_BY_COMMAND_ID,
  FILTER_MENU_ID,
  NUMBER_FORMAT_MENU_ID,
  PASTE_SPECIAL_OPEN_COMMAND_ID,
  PICK_FROM_LIST_COMMAND_ID,
  SHELL_COMMAND_ID,
  cellContextMenuLayout,
  headerContextMenuLayout,
} from '../src/renderer/cell-context-menu'
import { numberFormatCategories } from '../src/renderer/number-format'

describe('cellContextMenuLayout', () => {
  const nodes = cellContextMenuLayout()
  const byId = (id: string) => nodes.find((node) => node.id === id)

  it('keeps Excel order inside each group and uses unique ids', () => {
    const ids = nodes.flatMap((node) => [node.id, ...(node.children ?? []).map((c) => c.id)])
    expect(new Set(ids).size).toBe(ids.length)
    const others = nodes.filter((node) => node.group === ContextMenuGroup.OTHERS)
    expect(others.map((node) => node.id)).toEqual([
      'genoffice.ctx.new-comment',
      'genoffice.ctx.reply-comment',
      'genoffice.ctx.delete-comment',
      'genoffice.ctx.resolve-comment',
      NUMBER_FORMAT_MENU_ID,
      'genoffice.ctx.format-cells',
      PICK_FROM_LIST_COMMAND_ID,
      'genoffice.ctx.define-name',
      'genoffice.ctx.link',
      'genoffice.ctx.insert-function',
    ])
    expect(others.map((node) => node.order)).toEqual(
      [...others.map((n) => n.order)].sort((a, b) => a - b),
    )
  })

  it('Paste Special… is its own command so the dialog can register later', () => {
    const node = byId(PASTE_SPECIAL_OPEN_COMMAND_ID)
    expect(node?.group).toBe(ContextMenuGroup.FORMAT)
    expect(node?.commandId).toBe(PASTE_SPECIAL_OPEN_COMMAND_ID)
    expect(node?.children).toBeUndefined()
  })

  it('Filter ▸ sits before Sort ▸ with value / fill / font / clear entries', () => {
    const filter = byId(FILTER_MENU_ID)
    expect(filter?.group).toBe(ContextMenuGroup.DATA)
    expect(filter?.order).toBeLessThan(0)
    expect(filter?.children?.map((c) => c.params?.by ?? c.id)).toEqual([
      'value',
      'fill',
      'font',
      CLEAR_FILTER_MENU_ID,
    ])
    for (const child of filter?.children?.slice(0, 3) ?? []) {
      expect(child.commandId).toBe(FILTER_BY_COMMAND_ID)
    }
  })

  it('Number Format ▸ lists every ribbon category plus More…, each as a format: command', () => {
    const children = byId(NUMBER_FORMAT_MENU_ID)?.children ?? []
    const categories = numberFormatCategories()
    expect(children).toHaveLength(categories.length + 1)
    categories.forEach((category, index) => {
      expect(children[index]?.commandId).toBe(SHELL_COMMAND_ID)
      expect(children[index]?.params?.command).toBe(`format:${category.pattern}`)
    })
    expect(children.at(-1)?.params?.command).toBe('format-cells')
  })

  it('dialog entries route ribbon commands through the shell sink', () => {
    const commands = nodes
      .filter((node) => node.commandId === SHELL_COMMAND_ID)
      .map((node) => node.params?.command)
    expect(commands).toEqual([
      'visual:paste',
      'comment-new',
      'comment-reply',
      'comment-delete',
      'comment-resolve',
      'format-cells',
      'name-manager-open',
      'link-open',
      'insert-function-open',
    ])
  })
})

describe('headerContextMenuLayout', () => {
  it('adds only Format Cells… to the row/column header menus', () => {
    const nodes = headerContextMenuLayout()
    expect(nodes).toHaveLength(1)
    expect(nodes[0]?.group).toBe(ContextMenuGroup.LAYOUT)
    expect(nodes[0]?.params?.command).toBe('format-cells')
  })
})

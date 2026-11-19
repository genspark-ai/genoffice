import type { MenuItemConstructorOptions } from 'electron'
import { extname } from 'node:path'
import type { TabSummary } from '../shared/tabs-api'

export type TabMenuStringKey =
  | 'menuTabSave'
  | 'menuTabSaveAs'
  | 'menuTabPrint'
  | 'menuTabRename'
  | 'menuTabDuplicate'
  | 'menuTabSplitBeside'
  | 'menuTabCopyPath'
  | 'menuTabCopyName'
  | 'menuTabReveal'
  | 'menuTabOpenDefault'
  | 'menuTabMoveLeft'
  | 'menuTabMoveRight'
  | 'menuTabCloseOthers'
  | 'menuTabCloseRight'
  | 'menuTabCloseAll'
  | 'menuTabReopen'
  | 'menuTabPin'
  | 'menuTabUnpin'
  | 'menuOpenInNewWindow'
  | 'menuClose'

/** What each menu entry does; the shell supplies these, the menu only decides what is offered. */
export interface TabMenuActions {
  save(id: string): void
  saveAs(id: string): void
  print(id: string): void
  openInNewWindow(id: string): void
  splitBeside(id: string): void
  rename(id: string): void
  duplicate(path: string): void
  copyText(text: string): void
  reveal(path: string): void
  openDefault(path: string): void
  move(id: string, toIndex: number): void
  /** close these tabs one after another (a save prompt that is cancelled stops the rest) */
  closeTabs(ids: string[]): void
  reopenClosed(): void
  /** keep the tab's renderer alive however long it sits in the background */
  pin(id: string, pinned: boolean): void
}

export interface TabMenuInput {
  tab: TabSummary
  /** every tab in strip order, Home first */
  tabs: readonly TabSummary[]
  canDetach: boolean
  /** the file behind the tab is still on disk */
  fileExists: boolean
  /** something was closed that can be opened again */
  canReopen: boolean
  /** Translate through the shell's existing 21-locale main table. */
  tm(key: TabMenuStringKey): string
  actions: TabMenuActions
}

/** Executable, script, shortcut, and launcher extensions that should not be launched via default app */
const PROGRAM_EXTENSIONS = new Set([
  // Windows executables, scripts, shortcuts, and installers
  '.exe',
  '.com',
  '.scr',
  '.msi',
  '.bat',
  '.cmd',
  '.ps1',
  '.vbs',
  '.vbe',
  '.js',
  '.jse',
  '.wsf',
  '.hta',
  '.lnk',
  '.url',
  '.reg',
  '.cpl',
  '.pif',
  '.gadget',
  '.msc',
  '.inf',
  '.ins',
  '.isp',
  // macOS applications, scripts, workflows, and packages
  '.app',
  '.command',
  '.action',
  '.workflow',
  '.scpt',
  '.applescript',
  '.terminal',
  '.pkg',
  '.mpkg',
  // Linux launchers, binaries, and shell scripts
  '.desktop',
  '.sh',
  '.bash',
  '.zsh',
  '.csh',
  '.ksh',
  '.bin',
  '.run',
  '.out',
  // Java archives
  '.jar',
])

export function isProgramFile(path: string): boolean {
  return PROGRAM_EXTENSIONS.has(extname(path).toLowerCase())
}

export const baseName = (path: string): string => path.split(/[\\/]/).pop() ?? path

const PATH_LABEL_MAX = 80

/**
 * A path short enough for a menu line: a long one keeps its start (the drive and first folder) and
 * its end (the folder the file is in and the file), with an ellipsis between.
 */
export function pathLabel(path: string): string {
  if (path.length <= PATH_LABEL_MAX) return path
  const keepEnd = Math.floor((PATH_LABEL_MAX - 1) * 0.65)
  const keepStart = PATH_LABEL_MAX - 1 - keepEnd
  return `${path.slice(0, keepStart)}…${path.slice(path.length - keepEnd)}`
}

/**
 * The right-click menu of a document tab: what to do with the tab (move, split, close the others)
 * and with its file (rename, copy, duplicate, reveal). Entries that need a file on disk are shown
 * but off for an untitled document, and "close" ones only count tabs that can be closed.
 */
export function tabMenuTemplate(input: TabMenuInput): MenuItemConstructorOptions[] {
  const { tab, tabs, tm, actions } = input
  const index = tabs.findIndex((other) => other.id === tab.id)
  const file = tab.filePath
  const onDisk = !!file && input.fileExists
  const closable = (other: TabSummary): boolean => other.id !== 'home' && other.closable
  const others = tabs.filter((other) => other.id !== tab.id && closable(other))
  const right = tabs.slice(index + 1).filter(closable)
  const everything = tabs.filter(closable)
  const document = input.canDetach && tab.kind !== 'home' && tab.kind !== 'help'

  const template: MenuItemConstructorOptions[] = [
    // where the file is, at a glance (the menu is the place people look when they ask "which one is this?")
    ...(file ? ([{ label: pathLabel(file), enabled: false }, { type: 'separator' }] as const) : []),
    { id: 'save', label: tm('menuTabSave'), enabled: document, click: () => actions.save(tab.id) },
    {
      id: 'saveAs',
      label: tm('menuTabSaveAs'),
      enabled: document,
      click: () => actions.saveAs(tab.id),
    },
    {
      id: 'print',
      label: tm('menuTabPrint'),
      enabled: document,
      click: () => actions.print(tab.id),
    },
    { type: 'separator' },
    {
      id: 'rename',
      label: tm('menuTabRename'),
      enabled: !!file,
      click: () => actions.rename(tab.id),
    },
    {
      id: 'duplicate',
      label: tm('menuTabDuplicate'),
      enabled: onDisk,
      click: () => actions.duplicate(file!),
    },
    { type: 'separator' },
    {
      id: 'splitBeside',
      label: tm('menuTabSplitBeside'),
      enabled: input.canDetach,
      click: () => actions.splitBeside(tab.id),
    },
    {
      id: 'openInNewWindow',
      label: tm('menuOpenInNewWindow'),
      enabled: input.canDetach,
      click: () => actions.openInNewWindow(tab.id),
    },
    { type: 'separator' },
    {
      id: 'copyPath',
      label: tm('menuTabCopyPath'),
      enabled: !!file,
      click: () => actions.copyText(file!),
    },
    {
      id: 'copyName',
      label: tm('menuTabCopyName'),
      enabled: !!file,
      click: () => actions.copyText(baseName(file!)),
    },
    {
      id: 'reveal',
      label: tm('menuTabReveal'),
      enabled: onDisk,
      click: () => actions.reveal(file!),
    },
    {
      id: 'openDefault',
      label: tm('menuTabOpenDefault'),
      enabled: onDisk && !isProgramFile(file!),
      click: () => actions.openDefault(file!),
    },
    { type: 'separator' },
    {
      id: 'pin',
      label: tm(tab.pinned ? 'menuTabUnpin' : 'menuTabPin'),
      enabled: tab.kind !== 'help',
      click: () => actions.pin(tab.id, !tab.pinned),
    },
    {
      id: 'moveLeft',
      label: tm('menuTabMoveLeft'),
      enabled: index > 1,
      click: () => actions.move(tab.id, index - 1),
    },
    {
      id: 'moveRight',
      label: tm('menuTabMoveRight'),
      enabled: index >= 1 && index < tabs.length - 1,
      click: () => actions.move(tab.id, index + 1),
    },
    { type: 'separator' },
    {
      id: 'close',
      label: tm('menuClose'),
      enabled: closable(tab),
      click: () => actions.closeTabs([tab.id]),
    },
    {
      id: 'closeOthers',
      label: tm('menuTabCloseOthers'),
      enabled: others.length > 0,
      click: () => actions.closeTabs(others.map((other) => other.id)),
    },
    {
      id: 'closeRight',
      label: tm('menuTabCloseRight'),
      enabled: right.length > 0,
      click: () => actions.closeTabs(right.map((other) => other.id)),
    },
    {
      id: 'closeAll',
      label: tm('menuTabCloseAll'),
      enabled: everything.length > 0,
      click: () => actions.closeTabs(everything.map((other) => other.id)),
    },
    { type: 'separator' },
    {
      id: 'reopen',
      label: tm('menuTabReopen'),
      enabled: input.canReopen,
      click: () => actions.reopenClosed(),
    },
  ]

  return template
}

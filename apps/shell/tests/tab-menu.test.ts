import { describe, expect, it, vi } from 'vitest'
import type { MenuItemConstructorOptions } from 'electron'
import type { TabSummary } from '../src/shared/tabs-api'
import { mainMenuStrings, translateMainMenu, tm } from './helpers/main-menu-strings'
import {
  baseName,
  isProgramFile,
  pathLabel,
  tabMenuTemplate,
  type TabMenuActions,
  type TabMenuInput,
} from '../src/main/tab-menu'

const tab = (id: string, extra: Partial<TabSummary> = {}): TabSummary => ({
  id,
  kind: 'docs',
  title: id,
  closable: true,
  active: false,
  ...extra,
})
const home = tab('home', { kind: 'home', closable: false })

function actions(): TabMenuActions {
  return {
    save: vi.fn(),
    saveAs: vi.fn(),
    print: vi.fn(),
    openInNewWindow: vi.fn(),
    splitBeside: vi.fn(),
    rename: vi.fn(),
    duplicate: vi.fn(),
    copyText: vi.fn(),
    reveal: vi.fn(),
    openDefault: vi.fn(),
    move: vi.fn(),
    closeTabs: vi.fn(),
    reopenClosed: vi.fn(),
  }
}

function menu(
  current: TabSummary,
  tabs: TabSummary[],
  extra: Partial<TabMenuInput> = {},
): {
  items: Map<string, MenuItemConstructorOptions>
  actions: TabMenuActions
  raw: MenuItemConstructorOptions[]
} {
  const given = actions()
  const raw = tabMenuTemplate({
    tab: current,
    tabs,
    canDetach: true,
    fileExists: true,
    canReopen: true,
    tm,
    actions: given,
    ...extra,
  })
  const items = new Map<string, MenuItemConstructorOptions>()
  for (const item of raw) if (item.label) items.set(item.label, item)
  return { items, actions: given, raw }
}
const click = (item: MenuItemConstructorOptions | undefined): void =>
  (item!.click as unknown as () => void)()
const enabled = (item: MenuItemConstructorOptions | undefined): boolean => item!.enabled !== false

describe('the tab context menu', () => {
  const a = tab('a', { filePath: 'D:\\docs\\Hợp đồng.docx' })
  const b = tab('b')
  const c = tab('c', { filePath: 'D:\\docs\\bảng.xlsx' })
  const all = [home, a, b, c]

  it('offers file actions for a saved document and runs them on its path', () => {
    const { items, actions: given } = menu(a, all)
    for (const label of [
      'Rename',
      'Duplicate (copy the saved file)',
      'Copy path',
      'Copy file name',
      'Show in folder',
      'Open with the default app',
    ])
      expect(enabled(items.get(label)), label).toBe(true)

    click(items.get('Copy path'))
    click(items.get('Copy file name'))
    click(items.get('Duplicate (copy the saved file)'))
    click(items.get('Rename'))
    expect(given.copyText).toHaveBeenNthCalledWith(1, 'D:\\docs\\Hợp đồng.docx')
    expect(given.copyText).toHaveBeenNthCalledWith(2, 'Hợp đồng.docx')
    expect(given.duplicate).toHaveBeenCalledWith('D:\\docs\\Hợp đồng.docx')
    expect(given.rename).toHaveBeenCalledWith('a')
  })

  it('correctly handles baseName for both Windows and POSIX paths', () => {
    expect(baseName('C:\\Users\\User\\Documents\\Report.docx')).toBe('Report.docx')
    expect(baseName('/home/user/documents/report.pdf')).toBe('report.pdf')
    expect(baseName('C:/Mixed/Slashes/data.xlsx')).toBe('data.xlsx')
    expect(baseName('D:\\tài liệu\\Báo cáo quý 3.docx')).toBe('Báo cáo quý 3.docx')
  })

  it('saves, saves a copy, and prints the clicked untitled tab', () => {
    const { items, actions: given } = menu(b, all)
    for (const label of ['Save', 'Save As…', 'Print…']) {
      expect(enabled(items.get(label))).toBe(true)
      click(items.get(label))
    }
    expect(given.save).toHaveBeenCalledWith('b')
    expect(given.saveAs).toHaveBeenCalledWith('b')
    expect(given.print).toHaveBeenCalledWith('b')
    const help = menu(tab('help', { kind: 'help' }), all).items
    for (const label of ['Save', 'Save As…', 'Print…']) expect(enabled(help.get(label))).toBe(false)
  })

  it('switches off what needs a file for an untitled document, or one gone from the disk', () => {
    const untitled = menu(b, all).items
    for (const label of [
      'Rename',
      'Duplicate (copy the saved file)',
      'Copy path',
      'Copy file name',
    ])
      expect(enabled(untitled.get(label)), label).toBe(false)

    const gone = menu(a, all, { fileExists: false }).items
    expect(enabled(gone.get('Copy path'))).toBe(true)
    for (const label of [
      'Duplicate (copy the saved file)',
      'Show in folder',
      'Open with the default app',
    ])
      expect(enabled(gone.get(label)), label).toBe(false)
  })

  it('prevents opening executable and program files with the default application', () => {
    // Windows executables, scripts, shortcuts, and installers
    expect(isProgramFile('C:\\Tools\\setup.exe')).toBe(true)
    expect(isProgramFile('/usr/bin/script.bat')).toBe(true)
    expect(isProgramFile('cmd.cmd')).toBe(true)
    expect(isProgramFile('run.ps1')).toBe(true)
    expect(isProgramFile('run.vbs')).toBe(true)
    expect(isProgramFile('shortcut.lnk')).toBe(true)
    expect(isProgramFile('internet.url')).toBe(true)
    expect(isProgramFile('legacy.pif')).toBe(true)
    expect(isProgramFile('widget.gadget')).toBe(true)
    // macOS scripts, applications, and workflows
    expect(isProgramFile('/tmp/x/script.command')).toBe(true)
    expect(isProgramFile('Automator.action')).toBe(true)
    expect(isProgramFile('Auto.workflow')).toBe(true)
    expect(isProgramFile('Script.scpt')).toBe(true)
    expect(isProgramFile('Script.applescript')).toBe(true)
    // Linux launchers and shell scripts
    expect(isProgramFile('/usr/share/applications/app.desktop')).toBe(true)
    expect(isProgramFile('run.sh')).toBe(true)
    expect(isProgramFile('script.bash')).toBe(true)
    expect(isProgramFile('script.zsh')).toBe(true)
    expect(isProgramFile('binary.bin')).toBe(true)
    // Case insensitivity and double extension tricks
    expect(isProgramFile('SETUP.EXE')).toBe(true)
    expect(isProgramFile('LAUNCH.COMMAND')).toBe(true)
    expect(isProgramFile('report.pdf.exe')).toBe(true)
    expect(isProgramFile('summary.docx.command')).toBe(true)

    // Valid documents supported by GenOffice and external tools must NOT be blocked
    expect(isProgramFile('document.docx')).toBe(false)
    expect(isProgramFile('spreadsheet.xlsx')).toBe(false)
    expect(isProgramFile('presentation.pptx')).toBe(false)
    expect(isProgramFile('document.pdf')).toBe(false)
    expect(isProgramFile('notes.md')).toBe(false)
    expect(isProgramFile('page.html')).toBe(false)
    expect(isProgramFile('data.csv')).toBe(false)
    expect(isProgramFile('readme.txt')).toBe(false)
    expect(isProgramFile('Báo cáo tổng hợp.docx')).toBe(false)
    expect(isProgramFile('\u65e5\u672c\u8a9eテスト.xlsx')).toBe(false)

    const programTab = tab('prog', { filePath: 'D:\\Downloads\\malicious.exe' })
    const progMenu = menu(programTab, [home, programTab], { fileExists: true }).items
    expect(enabled(progMenu.get('Open with the default app'))).toBe(false)

    const commandTab = tab('cmd-tab', { filePath: '/tmp/x/script.command' })
    const cmdMenu = menu(commandTab, [home, commandTab], { fileExists: true }).items
    expect(enabled(cmdMenu.get('Open with the default app'))).toBe(false)

    const desktopTab = tab('dsk-tab', { filePath: '/home/user/app.desktop' })
    const dskMenu = menu(desktopTab, [home, desktopTab], { fileExists: true }).items
    expect(enabled(dskMenu.get('Open with the default app'))).toBe(false)
  })

  it('closes the other tabs, the tabs to the right, or all of them, never Home', () => {
    const { items, actions: given } = menu(a, all)
    click(items.get('Close'))
    click(items.get('Close other tabs'))
    click(items.get('Close tabs to the right'))
    click(items.get('Close all tabs'))
    expect(given.closeTabs).toHaveBeenNthCalledWith(1, ['a'])
    expect(given.closeTabs).toHaveBeenNthCalledWith(2, ['b', 'c'])
    expect(given.closeTabs).toHaveBeenNthCalledWith(3, ['b', 'c'])
    expect(given.closeTabs).toHaveBeenNthCalledWith(4, ['a', 'b', 'c'])
  })

  it('leaves out tabs that cannot be closed and switches off what has nothing to do', () => {
    const pinned = tab('pinned', { closable: false })
    const last = menu(c, [home, pinned, c]).items
    expect(enabled(last.get('Close tabs to the right'))).toBe(false)
    expect(enabled(last.get('Move right'))).toBe(false)
    expect(enabled(last.get('Move left'))).toBe(true)

    const only = menu(a, [home, a]).items
    expect(enabled(only.get('Close other tabs'))).toBe(false)
    expect(enabled(only.get('Close all tabs'))).toBe(true)
    expect(enabled(only.get('Move left'))).toBe(false)

    // the tab strip keeps Home first: the first document tab cannot move into its place
    expect(enabled(menu(a, all).items.get('Move left'))).toBe(false)
  })

  it('moves a tab one place and splits it into a window of its own', () => {
    const { items, actions: given } = menu(b, all)
    click(items.get('Move left'))
    click(items.get('Move right'))
    click(items.get('Split side by side'))
    click(items.get('Open in New Window'))
    expect(given.move).toHaveBeenNthCalledWith(1, 'b', 1)
    expect(given.move).toHaveBeenNthCalledWith(2, 'b', 3)
    expect(given.splitBeside).toHaveBeenCalledWith('b')
    expect(given.openInNewWindow).toHaveBeenCalledWith('b')

    const stuck = menu(b, all, { canDetach: false }).items
    expect(enabled(stuck.get('Split side by side'))).toBe(false)
    expect(enabled(stuck.get('Open in New Window'))).toBe(false)
  })

  it('disables document commands for the chrome-free Present tab', () => {
    const preview = tab('preview', { kind: 'html', filePath: '/tmp/x/page.html' })
    const { items } = menu(preview, [home, preview], { canDetach: false })
    for (const label of ['Save', 'Save As…', 'Print…'])
      expect(enabled(items.get(label))).toBe(false)
    expect(enabled(items.get('Close'))).toBe(true)
  })

  it('reopens a closed tab only when there is one', () => {
    expect(enabled(menu(a, all, { canReopen: false }).items.get('Reopen closed tab'))).toBe(false)
    const { items, actions: given } = menu(a, all)
    click(items.get('Reopen closed tab'))
    expect(given.reopenClosed).toHaveBeenCalled()
  })

  it('uses the main table for every supported locale without English fallback', () => {
    const locales = [
      'zh',
      'en',
      'vi',
      'ja',
      'ko',
      'fr',
      'de',
      'es',
      'th',
      'id',
      'ru',
      'ar',
      'pt',
      'it',
      'pl',
      'cs',
      'nl',
      'ms',
      'he',
      'hi',
      'zh-TW',
    ] as const
    expect(Object.keys(mainMenuStrings).sort()).toEqual([...locales].sort())
    for (const locale of locales) {
      const localized = menu(a, all, { tm: (key) => translateMainMenu(locale, key) }).raw
      for (const item of localized) {
        if (!item.id) continue
        const key =
          item.id === 'openInNewWindow'
            ? 'menuOpenInNewWindow'
            : item.id === 'close'
              ? 'menuClose'
              : `menuTab${item.id[0]!.toUpperCase()}${item.id.slice(1)}`
        expect(mainMenuStrings[locale]).toHaveProperty(key)
        expect(item.label).toBe(mainMenuStrings[locale][key as keyof typeof mainMenuStrings.en])
        expect(item.label?.trim().length).toBeGreaterThan(0)
      }
    }
    expect(translateMainMenu('vi', 'menuTabCloseOthers')).toBe('Đóng các thẻ khác')
    expect(translateMainMenu('zh', 'menuTabCloseOthers')).toBe(
      '\u5173\u95ed\u5176\u4ed6\u6807\u7b7e\u9875',
    )
    expect(translateMainMenu('th', 'menuTabSave')).toBe('บันทึก')
    expect(translateMainMenu('id', 'menuTabSave')).toBe('Simpan')
  })
})

describe('where the file is, shown in the tab menu', () => {
  const path = 'D:\\Documents\\Project\\report.pdf'
  const first = (template: MenuItemConstructorOptions[]) => template[0]!

  it('starts with the file path as a line that cannot be clicked', () => {
    const template = tabMenuTemplate({
      tab: tab('a', { filePath: path }),
      tabs: [home, tab('a', { filePath: path })],
      canDetach: true,
      fileExists: true,
      canReopen: false,
      tm,
      actions: actions(),
    })
    expect(first(template)).toMatchObject({ label: path, enabled: false })
    expect(template[1]).toMatchObject({ type: 'separator' })
  })

  it('has no path line for an untitled document', () => {
    const template = tabMenuTemplate({
      tab: tab('a'),
      tabs: [home, tab('a')],
      canDetach: true,
      fileExists: false,
      canReopen: false,
      tm,
      actions: actions(),
    })
    expect(first(template).label).toBe('Save')
  })

  it('keeps the start and the end of a very long path', () => {
    const long = `D:\\${'folder\\'.repeat(30)}final-report.pdf`
    const label = pathLabel(long)
    expect(label.length).toBeLessThanOrEqual(80)
    expect(label.startsWith('D:\\folder')).toBe(true)
    expect(label.endsWith('final-report.pdf')).toBe(true)
    expect(label).toContain('…')
    expect(pathLabel(path)).toBe(path)
  })
})

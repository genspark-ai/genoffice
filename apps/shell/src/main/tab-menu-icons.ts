import { readFileSync } from 'node:fs'
import { nativeImage, nativeTheme, type NativeImage } from 'electron'
export type TabMenuIcon =
  | 'save'
  | 'saveAs'
  | 'print'
  | 'rename'
  | 'duplicate'
  | 'splitBeside'
  | 'openInNewWindow'
  | 'copyPath'
  | 'copyName'
  | 'reveal'
  | 'openDefault'
  | 'moveLeft'
  | 'moveRight'
  | 'close'
  | 'closeOthers'
  | 'closeRight'
  | 'closeAll'
  | 'reopen'

// Light mode assets
import save1x from './assets/tab-actions/png/save.png?asset'
import save2x from './assets/tab-actions/png/save@2x.png?asset'
import saveAs1x from './assets/tab-actions/png/saveAs.png?asset'
import saveAs2x from './assets/tab-actions/png/saveAs@2x.png?asset'
import print1x from './assets/tab-actions/png/print.png?asset'
import print2x from './assets/tab-actions/png/print@2x.png?asset'
import rename1x from './assets/tab-actions/png/rename.png?asset'
import rename2x from './assets/tab-actions/png/rename@2x.png?asset'
import duplicate1x from './assets/tab-actions/png/duplicate.png?asset'
import duplicate2x from './assets/tab-actions/png/duplicate@2x.png?asset'
import splitBeside1x from './assets/tab-actions/png/splitBeside.png?asset'
import splitBeside2x from './assets/tab-actions/png/splitBeside@2x.png?asset'
import openInNewWindow1x from './assets/tab-actions/png/openInNewWindow.png?asset'
import openInNewWindow2x from './assets/tab-actions/png/openInNewWindow@2x.png?asset'
import copyPath1x from './assets/tab-actions/png/copyPath.png?asset'
import copyPath2x from './assets/tab-actions/png/copyPath@2x.png?asset'
import copyName1x from './assets/tab-actions/png/copyName.png?asset'
import copyName2x from './assets/tab-actions/png/copyName@2x.png?asset'
import reveal1x from './assets/tab-actions/png/reveal.png?asset'
import reveal2x from './assets/tab-actions/png/reveal@2x.png?asset'
import openDefault1x from './assets/tab-actions/png/openDefault.png?asset'
import openDefault2x from './assets/tab-actions/png/openDefault@2x.png?asset'
import moveLeft1x from './assets/tab-actions/png/moveLeft.png?asset'
import moveLeft2x from './assets/tab-actions/png/moveLeft@2x.png?asset'
import moveRight1x from './assets/tab-actions/png/moveRight.png?asset'
import moveRight2x from './assets/tab-actions/png/moveRight@2x.png?asset'
import close1x from './assets/tab-actions/png/close.png?asset'
import close2x from './assets/tab-actions/png/close@2x.png?asset'
import closeOthers1x from './assets/tab-actions/png/closeOthers.png?asset'
import closeOthers2x from './assets/tab-actions/png/closeOthers@2x.png?asset'
import closeRight1x from './assets/tab-actions/png/closeRight.png?asset'
import closeRight2x from './assets/tab-actions/png/closeRight@2x.png?asset'
import closeAll1x from './assets/tab-actions/png/closeAll.png?asset'
import closeAll2x from './assets/tab-actions/png/closeAll@2x.png?asset'
import reopen1x from './assets/tab-actions/png/reopen.png?asset'
import reopen2x from './assets/tab-actions/png/reopen@2x.png?asset'

// Dark mode assets
import saveDark1x from './assets/tab-actions/png/save-dark.png?asset'
import saveDark2x from './assets/tab-actions/png/save-dark@2x.png?asset'
import saveAsDark1x from './assets/tab-actions/png/saveAs-dark.png?asset'
import saveAsDark2x from './assets/tab-actions/png/saveAs-dark@2x.png?asset'
import printDark1x from './assets/tab-actions/png/print-dark.png?asset'
import printDark2x from './assets/tab-actions/png/print-dark@2x.png?asset'
import renameDark1x from './assets/tab-actions/png/rename-dark.png?asset'
import renameDark2x from './assets/tab-actions/png/rename-dark@2x.png?asset'
import duplicateDark1x from './assets/tab-actions/png/duplicate-dark.png?asset'
import duplicateDark2x from './assets/tab-actions/png/duplicate-dark@2x.png?asset'
import splitBesideDark1x from './assets/tab-actions/png/splitBeside-dark.png?asset'
import splitBesideDark2x from './assets/tab-actions/png/splitBeside-dark@2x.png?asset'
import openInNewWindowDark1x from './assets/tab-actions/png/openInNewWindow-dark.png?asset'
import openInNewWindowDark2x from './assets/tab-actions/png/openInNewWindow-dark@2x.png?asset'
import copyPathDark1x from './assets/tab-actions/png/copyPath-dark.png?asset'
import copyPathDark2x from './assets/tab-actions/png/copyPath-dark@2x.png?asset'
import copyNameDark1x from './assets/tab-actions/png/copyName-dark.png?asset'
import copyNameDark2x from './assets/tab-actions/png/copyName-dark@2x.png?asset'
import revealDark1x from './assets/tab-actions/png/reveal-dark.png?asset'
import revealDark2x from './assets/tab-actions/png/reveal-dark@2x.png?asset'
import openDefaultDark1x from './assets/tab-actions/png/openDefault-dark.png?asset'
import openDefaultDark2x from './assets/tab-actions/png/openDefault-dark@2x.png?asset'
import moveLeftDark1x from './assets/tab-actions/png/moveLeft-dark.png?asset'
import moveLeftDark2x from './assets/tab-actions/png/moveLeft-dark@2x.png?asset'
import moveRightDark1x from './assets/tab-actions/png/moveRight-dark.png?asset'
import moveRightDark2x from './assets/tab-actions/png/moveRight-dark@2x.png?asset'
import closeDark1x from './assets/tab-actions/png/close-dark.png?asset'
import closeDark2x from './assets/tab-actions/png/close-dark@2x.png?asset'
import closeOthersDark1x from './assets/tab-actions/png/closeOthers-dark.png?asset'
import closeOthersDark2x from './assets/tab-actions/png/closeOthers-dark@2x.png?asset'
import closeRightDark1x from './assets/tab-actions/png/closeRight-dark.png?asset'
import closeRightDark2x from './assets/tab-actions/png/closeRight-dark@2x.png?asset'
import closeAllDark1x from './assets/tab-actions/png/closeAll-dark.png?asset'
import closeAllDark2x from './assets/tab-actions/png/closeAll-dark@2x.png?asset'
import reopenDark1x from './assets/tab-actions/png/reopen-dark.png?asset'
import reopenDark2x from './assets/tab-actions/png/reopen-dark@2x.png?asset'

interface IconPair {
  path1x: string
  path2x: string
}

const LIGHT_ASSETS: Record<TabMenuIcon, IconPair> = {
  save: { path1x: save1x, path2x: save2x },
  saveAs: { path1x: saveAs1x, path2x: saveAs2x },
  print: { path1x: print1x, path2x: print2x },
  rename: { path1x: rename1x, path2x: rename2x },
  duplicate: { path1x: duplicate1x, path2x: duplicate2x },
  splitBeside: { path1x: splitBeside1x, path2x: splitBeside2x },
  openInNewWindow: { path1x: openInNewWindow1x, path2x: openInNewWindow2x },
  copyPath: { path1x: copyPath1x, path2x: copyPath2x },
  copyName: { path1x: copyName1x, path2x: copyName2x },
  reveal: { path1x: reveal1x, path2x: reveal2x },
  openDefault: { path1x: openDefault1x, path2x: openDefault2x },
  moveLeft: { path1x: moveLeft1x, path2x: moveLeft2x },
  moveRight: { path1x: moveRight1x, path2x: moveRight2x },
  close: { path1x: close1x, path2x: close2x },
  closeOthers: { path1x: closeOthers1x, path2x: closeOthers2x },
  closeRight: { path1x: closeRight1x, path2x: closeRight2x },
  closeAll: { path1x: closeAll1x, path2x: closeAll2x },
  reopen: { path1x: reopen1x, path2x: reopen2x },
}

const DARK_ASSETS: Record<TabMenuIcon, IconPair> = {
  save: { path1x: saveDark1x, path2x: saveDark2x },
  saveAs: { path1x: saveAsDark1x, path2x: saveAsDark2x },
  print: { path1x: printDark1x, path2x: printDark2x },
  rename: { path1x: renameDark1x, path2x: renameDark2x },
  duplicate: { path1x: duplicateDark1x, path2x: duplicateDark2x },
  splitBeside: { path1x: splitBesideDark1x, path2x: splitBesideDark2x },
  openInNewWindow: { path1x: openInNewWindowDark1x, path2x: openInNewWindowDark2x },
  copyPath: { path1x: copyPathDark1x, path2x: copyPathDark2x },
  copyName: { path1x: copyNameDark1x, path2x: copyNameDark2x },
  reveal: { path1x: revealDark1x, path2x: revealDark2x },
  openDefault: { path1x: openDefaultDark1x, path2x: openDefaultDark2x },
  moveLeft: { path1x: moveLeftDark1x, path2x: moveLeftDark2x },
  moveRight: { path1x: moveRightDark1x, path2x: moveRightDark2x },
  close: { path1x: closeDark1x, path2x: closeDark2x },
  closeOthers: { path1x: closeOthersDark1x, path2x: closeOthersDark2x },
  closeRight: { path1x: closeRightDark1x, path2x: closeRightDark2x },
  closeAll: { path1x: closeAllDark1x, path2x: closeAllDark2x },
  reopen: { path1x: reopenDark1x, path2x: reopenDark2x },
}

function loadActionIcon(pair: IconPair): NativeImage | null {
  try {
    const icon = nativeImage.createFromPath(pair.path1x)
    if (icon.isEmpty()) return null
    try {
      const buffer2x = readFileSync(pair.path2x)
      icon.addRepresentation({ scaleFactor: 2, buffer: buffer2x })
    } catch {
      // 1x representation remains valid if 2x fails to read
    }
    if (process.platform === 'darwin') {
      icon.setTemplateImage(true)
    }
    return icon
  } catch {
    return null
  }
}

const iconCache = new Map<boolean, Partial<Record<TabMenuIcon, NativeImage>>>()

/**
 * Returns the cached native action icons for the tab context menu,
 * appropriately styled for the current OS and light/dark theme.
 */
export function tabActionIcons(isDarkTheme?: boolean): Partial<Record<TabMenuIcon, NativeImage>> {
  const dark =
    typeof isDarkTheme === 'boolean'
      ? isDarkTheme
      : process.platform !== 'darwin' && nativeTheme.shouldUseDarkColors

  const cached = iconCache.get(dark)
  if (cached) return cached

  const assetMap = dark ? DARK_ASSETS : LIGHT_ASSETS
  const result: Partial<Record<TabMenuIcon, NativeImage>> = {}

  for (const [key, pair] of Object.entries(assetMap)) {
    const icon = loadActionIcon(pair)
    if (icon && !icon.isEmpty()) {
      result[key as TabMenuIcon] = icon
    }
  }

  iconCache.set(dark, result)
  return result
}

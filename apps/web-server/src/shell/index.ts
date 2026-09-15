/**
 * Shell module entry — wires every shell sub-domain (app info, home,
 * tabs, devices, search, speech, notifications, cloud, offline, charts,
 * files, clipboard, windows) into the shared registry.
 */
import { registerAppInfoHandlers } from './app-info'
import { registerChartHandlers } from './charts'
import { registerClipboardHandlers } from './clipboard'
import { registerCloudHandlers } from './cloud'
import { registerMobileHandlers, registerMultimodalHandlers } from './devices'
import { registerFilesHandlers } from './files'
import { registerHomeHandlers } from './home'
import { registerModuleHandlers } from './modules'
import { registerSkillHandlers } from './skills'
import { registerNotificationHandlers } from './notifications'
import { registerPrefsHandlers } from './prefs'
import { registerOfflineHandlers } from './offline'
import { registerSearchHandlers } from './search'
import { registerSpeechHandlers } from './speech'
import { registerTabsHandlers, registerUpdateHandlers } from './tabs'
import { registerWindowHandlers } from './windows'

export function registerShellHandlers(): void {
  registerAppInfoHandlers()
  registerHomeHandlers()
  registerModuleHandlers()
  registerSkillHandlers()
  registerTabsHandlers()
  registerUpdateHandlers()
  registerWindowHandlers()
  registerClipboardHandlers()
  registerFilesHandlers()
  registerNotificationHandlers()
  registerCloudHandlers()
  registerOfflineHandlers()
  registerMobileHandlers()
  registerMultimodalHandlers()
  registerSearchHandlers()
  registerSpeechHandlers()
  registerChartHandlers()
  registerPrefsHandlers()
}

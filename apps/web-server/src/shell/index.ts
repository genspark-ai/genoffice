/**
 * Shell module entry — wires every shell sub-domain (app info, home,
 * tabs, devices, search, speech, notifications, cloud, offline, charts,
 * files, clipboard, windows) into the shared registry.
 */
import { registerAppInfoHandlers } from './app-info.js'
import { registerChartHandlers } from './charts.js'
import { registerClipboardHandlers } from './clipboard.js'
import { registerCloudHandlers } from './cloud.js'
import { registerMobileHandlers, registerMultimodalHandlers } from './devices.js'
import { registerFilesHandlers } from './files.js'
import { registerHomeHandlers } from './home.js'
import { registerNotificationHandlers } from './notifications.js'
import { registerOfflineHandlers } from './offline.js'
import { registerSearchHandlers } from './search.js'
import { registerSpeechHandlers } from './speech.js'
import { registerTabsHandlers, registerUpdateHandlers } from './tabs.js'
import { registerWindowHandlers } from './windows.js'

export function registerShellHandlers(): void {
  registerAppInfoHandlers()
  registerHomeHandlers()
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
}

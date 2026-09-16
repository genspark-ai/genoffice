/**
 * Shell module entry — wires every shell sub-domain (app info, home,
 * tabs, devices, search, speech, notifications, cloud, offline, charts,
 * files, clipboard, windows) into the shared registry.
 */
import { registerAppInfoHandlers } from './app-info'
import { registerChartHandlers } from './charts'
import { ensureLumosSkillsRegistered, ensureSkillDirRegistered, ensureBuiltInSkillsMaterialized } from './pi-resources'
import { registerPiSessionHandlers } from './pi-session'
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
  // Materialise pi skill wrappers BEFORE any IPC handler that may try to install a
  // skill — otherwise the first install can race the wrapper write and the SKILL.md
  // for the new id is invisible to pi's loader on the next session.
  bootstrapPiSkills()
  // Lazy-build the embedded pi AgentSession and expose its tools/skills through
  // the standard IPC. The first call to `home:pi-status` triggers the actual
  // session creation so server startup stays snappy.
  registerPiSessionHandlers()
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


/**
 * Run the pi-resource bootstrap once at server startup.
 *
 *  1. Materialise SKILL.md wrappers for every LumosAI bundled skill so the pi
 *     agent can `bash` into the upstream Python handlers (the only path that
 *     preserves the mature LumosAI scripts without re-implementing them in TS).
 *  2. Point pi's settings at our wrapper root + the marketplace skills dir so
 *     the agent sees them on the next `createAgentSession()`.
 *  3. Materialise SKILL.md files for built-in skills so marketplace-installed
 *     skills have real pi-side presence (the marketplace catalog is metadata
 *     only — there is no on-disk SKILL.md until we write one here).
 *
 * Errors are logged, not thrown — a failed bootstrap must not block the rest of
 * the shell from wiring up.
 */
function bootstrapPiSkills(): void {
  void (async () => {
    try {
      const builtIn = ensureBuiltInSkillsMaterialized()
      if (builtIn.written.length || builtIn.skipped.length) {
        console.log(
          `[shell] materialised ${builtIn.written.length} built-in skills` +
          (builtIn.skipped.length ? ` (${builtIn.skipped.length} already on disk)` : ''),
        )
      }
      const lumos = await ensureLumosSkillsRegistered()
      if (lumos.registered.length || lumos.alreadyHad.length) {
        console.log(
          `[shell] lumos skill wrappers: ${lumos.registered.length} new, ${lumos.alreadyHad.length} cached`,
        )
      }
      await ensureSkillDirRegistered()
    } catch (err) {
      console.warn('[shell] pi skill bootstrap failed:', err)
    }
  })()
}

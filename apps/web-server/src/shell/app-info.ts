/**
 * Shell app info channels — language, version, platform, theme. These
 * mirror `apps/shell/src/main/index.ts` semantics for the standalone web
 * server build.
 */
import { registerHandle } from '../common/index.js'

export function registerAppInfoHandlers(): void {
  registerHandle('app:get-language', () => 'zh-CN')
  registerHandle('app:get-version', () => '0.8.0')
  registerHandle('app:get-platform', () => 'web')
  registerHandle('app:get-theme', () => ({
    theme: 'system',
    darkMode: false,
    highContrast: false,
  }))
}

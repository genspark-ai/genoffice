import { shell } from 'electron'
import { CodexAuthService } from '@genoffice/ai-provider'
import { beginCodexCallback, codexCredentialStore } from '@genoffice/ai-provider/codex-auth-node'
import { safeExternalUrl } from '@genoffice/electron-utils'

let codexAuth: CodexAuthService | undefined

export function getCodexAuth(): CodexAuthService {
  codexAuth ??= new CodexAuthService({
    store: codexCredentialStore(),
    clock: () => Date.now(),
    fetch,
    openBrowser: async (url) => {
      const target = safeExternalUrl(url)
      if (!target) throw new Error('ChatGPT sign-in URL rejected')
      await shell.openExternal(target)
    },
    beginCallback: beginCodexCallback,
  })
  return codexAuth
}

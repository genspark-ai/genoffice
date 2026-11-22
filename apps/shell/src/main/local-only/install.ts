import { join } from 'node:path'
import { app, dialog, session, shell } from 'electron'
import {
  createBlockLog,
  installElectronEgressGuard,
  installNodeEgressGuard,
  pinChildProxyEnv,
  type ElectronEgressDeps,
} from './egress-guard'

/**
 * Local-only fork: installs the egress guard as a side effect of being
 * imported. `apps/shell/src/main/index.ts` imports this module first, so the
 * guard is in place before any other module runs. See docs/local-only-fork.md.
 */

// resolved per write: the userData override in index.ts runs after this module
const report = createBlockLog(() => join(app.getPath('userData'), 'egress-blocked.log'))

installNodeEgressGuard(report)
pinChildProxyEnv(process.env)
installElectronEgressGuard({ app, session, shell, dialog } as unknown as ElectronEgressDeps, report)

/**
 * React bindings for the office session.
 *
 * PiSessionProvider owns the session lifetime; children use hooks to subscribe to
 * events, dialogs, notifications, and statuses.
 *
 * Designed so that:
 * - The session is created exactly once per provider (and disposed on unmount)
 * - The same uiAdapter is shared between the session and React
 * - Children can re-render in response to dialogs/notifications without forcing
 *   the entire tree to re-render (each hook subscribes independently)
 */

import {
  createContext,
  useContext,
  useEffect,
  useState,
  useSyncExternalStore,
  type ReactNode,
} from "react";
import { ReactUIAdapter, type DialogRequest, type NotificationItem } from "./ui-adapter";
import { createOfficeSession, type OfficeSession, type OfficeSessionOptions } from "./session";

/** OfficeSessionContext is exported for tests; production code should use
 *  the hooks (useOfficeSession, usePiDialogs, etc.) instead. */
export const OfficeSessionContext = createContext<OfficeSession | null>(null);

export interface PiSessionProviderProps extends OfficeSessionOptions {
  children: ReactNode;
  /**
   * If false, the provider will not create a session — useful for tests or
   * when the session is provided externally via `value`.
   */
  create?: boolean;
  /** Externally-created session (skips internal createOfficeSession). */
  value?: OfficeSession | null;
}

export function PiSessionProvider(props: PiSessionProviderProps) {
  const { children, create = true, value, ...opts } = props;
  const [session, setSession] = useState<OfficeSession | null>(value ?? null);
  const [error, setError] = useState<Error | null>(null);

  useEffect(() => {
    if (value) {
      setSession(value);
      return;
    }
    if (!create) return;
    let cancelled = false;
    let created: OfficeSession | null = null;
    createOfficeSession(opts)
      .then((s) => {
        if (cancelled) {
          s.dispose();
          return;
        }
        created = s;
        setSession(s);
      })
      .catch((e) => {
        if (!cancelled) setError(e instanceof Error ? e : new Error(String(e)));
      });
    return () => {
      cancelled = true;
      if (created) created.dispose();
      setSession(null);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [opts.cwd, opts.agentDir, value, create]);

  if (error) throw error;
  return <OfficeSessionContext.Provider value={session}>{children}</OfficeSessionContext.Provider>;
}

/** Get the full OfficeSession (session + uiAdapter). Throws if not in a provider. */
export function useOfficeSession(): OfficeSession {
  const ctx = useContext(OfficeSessionContext);
  if (!ctx) throw new Error("useOfficeSession must be used inside <PiSessionProvider>");
  return ctx;
}

/** Back-compat alias matching the plan's name. */
export function usePiSession(): OfficeSession {
  return useOfficeSession();
}

/** Get just the pi AgentSession. */
export function usePiAgentSession() {
  return useOfficeSession().session;
}

/** Get just the ReactUIAdapter. */
export function useUiAdapter(): ReactUIAdapter {
  return useOfficeSession().uiAdapter;
}

// ---------------------------------------------------------------------------
// Subscriptions via useSyncExternalStore (concurrent-safe, no re-render churn)
// ---------------------------------------------------------------------------

function subscribeToAdapter<T>(subscribe: (cb: () => void) => () => void, get: () => T): T {
  return useSyncExternalStore(subscribe, get, get);
}

export function usePiDialogs(): readonly DialogRequest[] {
  const adapter = useUiAdapter();
  return subscribeToAdapter(
    (cb) => adapter.onDialogs(() => cb()),
    () => adapter.dialogs,
  );
}

export function usePiNotifications(): readonly NotificationItem[] {
  const adapter = useUiAdapter();
  return subscribeToAdapter(
    (cb) => adapter.onNotifications(() => cb()),
    () => adapter.notifications,
  );
}

export function usePiStatuses(): ReadonlyMap<string, string | undefined> {
  const adapter = useUiAdapter();
  return subscribeToAdapter(
    (cb) => adapter.onStatuses(() => cb()),
    () => adapter.statuses,
  );
}

// Re-export the adapter for components that need to resolve dialogs directly.
export { ReactUIAdapter };

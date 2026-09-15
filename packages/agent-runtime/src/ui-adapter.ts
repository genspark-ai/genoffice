/**
 * ReactUIAdapter — bridges pi's ExtensionUIContext with React state.
 *
 * The adapter is a framework-agnostic event store:
 * - Tools call ctx.ui.confirm(...) etc., which push a DialogRequest into a queue
 * - React components subscribe via usePiDialogs/usePiNotifications and render modals
 * - When the user picks an option, the adapter resolves the promise and removes the request
 *
 * This keeps the runtime package free of React imports in the core, while still
 * allowing React UI to drive the dialog flow.
 */

import type { ExtensionUIContext, ExtensionUIDialogOptions } from "@earendil-works/pi-coding-agent";

// ============================================================================
// Event shapes
// ============================================================================

export type DialogKind = "select" | "confirm" | "input";

export interface DialogRequestBase {
  id: number;
  kind: DialogKind;
  title: string;
  signal?: AbortSignal;
  timeout?: number;
}

export interface SelectDialogRequest extends DialogRequestBase {
  kind: "select";
  options: string[];
  resolve: (value: string | undefined) => void;
}

export interface ConfirmDialogRequest extends DialogRequestBase {
  kind: "confirm";
  message: string;
  resolve: (value: boolean) => void;
}

export interface InputDialogRequest extends DialogRequestBase {
  kind: "input";
  placeholder?: string;
  resolve: (value: string | undefined) => void;
}

export type DialogRequest = SelectDialogRequest | ConfirmDialogRequest | InputDialogRequest;

export type NotificationKind = "info" | "warning" | "error";

export interface NotificationItem {
  id: number;
  message: string;
  type: NotificationKind;
}

// ============================================================================
// Listeners
// ============================================================================

export type DialogListener = (dialogs: readonly DialogRequest[]) => void;
export type NotificationListener = (items: readonly NotificationItem[]) => void;
export type StatusListener = (statuses: ReadonlyMap<string, string | undefined>) => void;

// ============================================================================
// ReactUIAdapter
// ============================================================================

export interface ReactUIAdapterOptions {
  /** Default timeout for dialogs in ms; if the user doesn't respond in time, the
   *  promise resolves with `undefined` / `false` (matching the noOpUIContext contract). */
  defaultDialogTimeoutMs?: number;
}

export class ReactUIAdapter implements ExtensionUIContext {
  private dialogsInternal: DialogRequest[] = [];
  private notificationsInternal: NotificationItem[] = [];
  private statusesInternal = new Map<string, string | undefined>();

  private dialogListeners = new Set<DialogListener>();
  private notificationListeners = new Set<NotificationListener>();
  private statusListeners = new Set<StatusListener>();

  private customData = new Map<string, unknown>();
  private nextId = 1;
  private nextNotificationId = 1;

  private defaultDialogTimeoutMs: number;

  // The pi runner holds a reference to the ExtensionUIContext. We expose a
  // `piContext` getter that returns this object (which implements the interface).
  // We don't need to do anything special — `this` IS the context.
  readonly piContext: ExtensionUIContext = this;

  constructor(opts: ReactUIAdapterOptions = {}) {
    this.defaultDialogTimeoutMs = opts.defaultDialogTimeoutMs ?? 5 * 60_000;
  }

  // --------------------------------------------------------------------------
  // React subscription API
  // --------------------------------------------------------------------------

  /** Snapshot of currently pending dialogs (read-only view). */
  get dialogs(): readonly DialogRequest[] {
    return this.dialogsInternal;
  }

  get notifications(): readonly NotificationItem[] {
    return this.notificationsInternal;
  }

  get statuses(): ReadonlyMap<string, string | undefined> {
    return this.statusesInternal;
  }

  onDialogs(listener: DialogListener): () => void {
    this.dialogListeners.add(listener);
    return () => this.dialogListeners.delete(listener);
  }

  onNotifications(listener: NotificationListener): () => void {
    this.notificationListeners.add(listener);
    return () => this.notificationListeners.delete(listener);
  }

  onStatuses(listener: StatusListener): () => void {
    this.statusListeners.add(listener);
    return () => this.statusListeners.delete(listener);
  }

  /** Resolve a pending dialog (called by React when user picks an option). */
  resolveDialog(id: number, value: string | boolean | undefined): void {
    const idx = this.dialogsInternal.findIndex((d) => d.id === id);
    if (idx === -1) return;
    const dlg = this.dialogsInternal[idx]!;
    this.dialogsInternal = this.dialogsInternal.filter((d) => d.id !== id);
    this.notifyDialogs();
    if (dlg.kind === "confirm") {
      dlg.resolve(Boolean(value));
    } else {
      dlg.resolve(typeof value === "string" ? value : undefined);
    }
  }

  /** Dismiss a notification (called by React after a TTL or user click). */
  dismissNotification(id: number): void {
    const idx = this.notificationsInternal.findIndex((n) => n.id === id);
    if (idx === -1) return;
    this.notificationsInternal = this.notificationsInternal.filter((n) => n.id !== id);
    this.notifyNotifications();
  }

  // --------------------------------------------------------------------------
  // Custom data (Office-specific bag for editor instance, frozen selection, etc.)
  // --------------------------------------------------------------------------

  setCustomData<T>(key: string, value: T): void {
    this.customData.set(key, value);
  }

  getCustomData<T>(key: string): T | undefined {
    return this.customData.get(key) as T | undefined;
  }

  deleteCustomData(key: string): boolean {
    return this.customData.delete(key);
  }

  // --------------------------------------------------------------------------
  // Editor instance (Office-specific convenience)
  // --------------------------------------------------------------------------

  setEditorInstance(editor: unknown): void {
    this.setCustomData("editor", editor);
  }

  getEditorInstance<T = unknown>(): T | undefined {
    return this.getCustomData<T>("editor");
  }

  // --------------------------------------------------------------------------
  // ExtensionUIContext implementation
  // --------------------------------------------------------------------------

  select(title: string, options: string[], opts?: ExtensionUIDialogOptions): Promise<string | undefined> {
    return this.pushDialog<string | undefined>({
      kind: "select",
      title,
      options,
      signal: opts?.signal,
      timeout: opts?.timeout,
    });
  }

  confirm(title: string, message: string, opts?: ExtensionUIDialogOptions): Promise<boolean> {
    return this.pushDialog<boolean>({
      kind: "confirm",
      title,
      message,
      signal: opts?.signal,
      timeout: opts?.timeout,
    });
  }

  input(title: string, placeholder?: string, opts?: ExtensionUIDialogOptions): Promise<string | undefined> {
    return this.pushDialog<string | undefined>({
      kind: "input",
      title,
      placeholder,
      signal: opts?.signal,
      timeout: opts?.timeout,
    });
  }

  notify(message: string, type: NotificationKind = "info"): void {
    const item: NotificationItem = { id: this.nextNotificationId++, message, type };
    this.notificationsInternal = [...this.notificationsInternal, item];
    this.notifyNotifications();
  }

  // Methods below exist for interface compliance but are not used by Office apps.
  // They match noOpUIContext so a non-React host can still create one.
  onTerminalInput(): () => void {
    return () => {};
  }
  setStatus(key: string, text: string | undefined): void {
    // Immutable Map update so React's useSyncExternalStore sees a fresh reference.
    const next = new Map(this.statusesInternal);
    if (text === undefined) {
      next.delete(key);
    } else {
      next.set(key, text);
    }
    this.statusesInternal = next;
    this.notifyStatuses();
  }
  setWorkingMessage(): void {}
  setWorkingVisible(): void {}
  setWorkingIndicator(): void {}
  setHiddenThinkingLabel(): void {}
  setWidget(): void {}
  setFooter(): void {}
  setHeader(): void {}
  setTitle(): void {}
  custom<T>(): Promise<T> {
    return Promise.resolve(undefined as T);
  }
  pasteToEditor(): void {}
  setEditorText(): void {}
  getEditorText(): string {
    return "";
  }
  editor(title: string, prefill?: string): Promise<string | undefined> {
    return this.input(title, prefill);
  }
  addAutocompleteProvider(): void {}
  setEditorComponent(): void {}
  getEditorComponent(): undefined {
    return undefined;
  }
  get theme(): never {
    throw new Error("ReactUIAdapter: theme not available in non-TUI mode");
  }
  getAllThemes(): Array<{ name: string; path: string | undefined }> {
    return [];
  }
  getTheme(): undefined {
    return undefined;
  }
  setTheme(): { success: boolean; error?: string } {
    return { success: false, error: "Themes are not supported in GenOffice runtime" };
  }
  getToolsExpanded(): boolean {
    return false;
  }
  setToolsExpanded(): void {}

  // --------------------------------------------------------------------------
  // Internals
  // --------------------------------------------------------------------------

  private pushDialog<T>(init: {
    kind: "select";
    title: string;
    options: string[];
    signal?: AbortSignal;
    timeout?: number;
  } | {
    kind: "confirm";
    title: string;
    message: string;
    signal?: AbortSignal;
    timeout?: number;
  } | {
    kind: "input";
    title: string;
    placeholder?: string;
    signal?: AbortSignal;
    timeout?: number;
  }): Promise<T> {
    return new Promise<T>((resolveOuter) => {
      const id = this.nextId++;
      const timeoutMs = init.timeout ?? this.defaultDialogTimeoutMs;
      let settled = false;

      const settle = (value: T | boolean | string | undefined) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        const stillThere = this.dialogsInternal.some((d) => d.id === id);
        if (stillThere) {
          this.dialogsInternal = this.dialogsInternal.filter((d) => d.id !== id);
          this.notifyDialogs();
        }
        resolveOuter(value as T);
      };

      const timer = setTimeout(() => {
        const defaultValue: boolean | string | undefined =
          init.kind === "confirm" ? false : undefined;
        settle(defaultValue);
      }, timeoutMs);

      const resolve: DialogRequest["resolve"] = ((v: unknown) => settle(v as never)) as DialogRequest["resolve"];

      let request: DialogRequest;
      if (init.kind === "select") {
        request = { id, kind: "select", title: init.title, options: init.options, resolve: resolve as SelectDialogRequest["resolve"] };
      } else if (init.kind === "confirm") {
        request = { id, kind: "confirm", title: init.title, message: init.message, resolve: resolve as ConfirmDialogRequest["resolve"] };
      } else {
        request = {
          id,
          kind: "input",
          title: init.title,
          ...(init.placeholder !== undefined ? { placeholder: init.placeholder } : {}),
          resolve: resolve as InputDialogRequest["resolve"],
        };
      }

      this.dialogsInternal = [...this.dialogsInternal, request];
      this.notifyDialogs();
    });
  }

  private notifyDialogs(): void {
    // Hand listeners a fresh reference so React's `===` comparison sees the change.
    const snapshot = this.dialogsInternal;
    for (const l of this.dialogListeners) l(snapshot);
  }
  private notifyNotifications(): void {
    const snapshot = this.notificationsInternal;
    for (const l of this.notificationListeners) l(snapshot);
  }
  private notifyStatuses(): void {
    for (const l of this.statusListeners) l(this.statusesInternal);
  }
}

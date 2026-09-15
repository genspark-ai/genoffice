/**
 * React UI components that bridge pi tool dialogs → user-visible modals.
 *
 * The intended integration:
 *   - The host app wraps its tree in <PiSessionProvider>
 *   - Somewhere near the root, render <PiDialogHost /> and <NotificationToaster />
 *   - When a tool calls ctx.ui.confirm(...), the dialog appears as a modal
 *   - When the user clicks OK/Cancel, adapter.resolveDialog(id, value) is called
 *   - The tool's promise resolves and the LLM sees the result
 *
 * These components are intentionally framework-light: no CSS framework, no
 * portals, no animation library. The host app can swap in its own modal
 * components by re-implementing the same hook subscriptions.
 */

import { useCallback, useState } from "react";
import { type DialogRequest, type NotificationItem } from "./ui-adapter";
import {
  usePiDialogs,
  usePiNotifications,
  usePiStatuses,
  useUiAdapter,
} from "./provider";

// ============================================================================
// PiDialogHost
// ============================================================================

/**
 * Renders all pending dialogs. Designed to be placed near the root of the
 * React tree so it floats above the rest of the UI.
 *
 * Visual style: a centered card with the title, body, and Cancel/OK buttons.
 * No external CSS — host apps can override via the optional `classNames` prop.
 */
export interface PiDialogHostProps {
  /** Optional CSS class names per element. */
  classNames?: Partial<{
    backdrop: string;
    card: string;
    title: string;
    message: string;
    actions: string;
    button: string;
    buttonPrimary: string;
    buttonDanger: string;
  }>;
  /** Show a countdown when a dialog has a timeout (default true). */
  showCountdown?: boolean;
}

const DEFAULT_CLASSNAMES = {
  backdrop: "pi-dialog-backdrop",
  card: "pi-dialog-card",
  title: "pi-dialog-title",
  message: "pi-dialog-message",
  actions: "pi-dialog-actions",
  button: "pi-dialog-btn",
  buttonPrimary: "pi-dialog-btn pi-dialog-btn--primary",
  buttonDanger: "pi-dialog-btn pi-dialog-btn--danger",
};

export function PiDialogHost(props: PiDialogHostProps = {}) {
  const dialogs = usePiDialogs();
  const adapter = useUiAdapter();
  const cn = { ...DEFAULT_CLASSNAMES, ...(props.classNames ?? {}) };

  if (dialogs.length === 0) return null;
  // Render the most recent dialog on top (last in queue).
  const dialog = dialogs[dialogs.length - 1]!;

  return (
    <PiDialogRenderer
      dialog={dialog}
      classNames={cn}
      onResolve={(value) => adapter.resolveDialog(dialog.id, value)}
      showCountdown={props.showCountdown !== false}
    />
  );
}

function PiDialogRenderer(props: {
  dialog: DialogRequest;
  classNames: typeof DEFAULT_CLASSNAMES;
  onResolve: (value: string | boolean | undefined) => void;
  showCountdown: boolean;
}) {
  const { dialog, classNames: cn, onResolve, showCountdown } = props;
  const [countdown, setCountdown] = useState<number | null>(null);

  // Optional countdown display
  useCountdown(dialog.timeout, showCountdown, setCountdown);

  const handleCancel = useCallback(() => onResolve(undefined), [onResolve]);

  if (dialog.kind === "confirm") {
    return (
      <div className={cn.backdrop} data-testid="pi-dialog-backdrop">
        <div className={cn.card} role="dialog" aria-labelledby="pi-dialog-title" aria-modal="true">
          <h2 id="pi-dialog-title" className={cn.title}>{dialog.title}</h2>
          <p className={cn.message}>{dialog.message}</p>
          {countdown !== null && (
            <p className={cn.message} data-testid="pi-dialog-countdown">
              Auto-cancelling in {countdown}s
            </p>
          )}
          <div className={cn.actions}>
            <button className={cn.button} onClick={() => onResolve(false)} data-testid="pi-dialog-cancel">
              Cancel
            </button>
            <button className={cn.buttonPrimary} onClick={() => onResolve(true)} data-testid="pi-dialog-ok">
              OK
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (dialog.kind === "input") {
    return (
      <InputDialog
        dialog={dialog}
        classNames={cn}
        onResolve={onResolve}
        countdown={countdown}
        onCancel={handleCancel}
      />
    );
  }

  // select
  return (
    <SelectDialog
      dialog={dialog}
      classNames={cn}
      onResolve={onResolve}
      countdown={countdown}
      onCancel={handleCancel}
    />
  );
}

function InputDialog(props: {
  dialog: Extract<DialogRequest, { kind: "input" }>;
  classNames: typeof DEFAULT_CLASSNAMES;
  onResolve: (v: string | undefined) => void;
  countdown: number | null;
  onCancel: () => void;
}) {
  const { dialog, classNames: cn, onResolve, countdown, onCancel } = props;
  const [value, setValue] = useState("");
  const submit = () => onResolve(value);
  return (
    <div className={cn.backdrop} data-testid="pi-dialog-backdrop">
      <div className={cn.card} role="dialog" aria-modal="true">
        <h2 className={cn.title}>{dialog.title}</h2>
        <input
          type="text"
          placeholder={dialog.placeholder ?? ""}
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") submit()
            if (e.key === "Escape") onCancel()
          }}
          className={cn.message}
          autoFocus
          data-testid="pi-dialog-input"
        />
        {countdown !== null && <p className={cn.message} data-testid="pi-dialog-countdown">Auto-cancelling in {countdown}s</p>}
        <div className={cn.actions}>
          <button className={cn.button} onClick={onCancel} data-testid="pi-dialog-cancel">Cancel</button>
          <button className={cn.buttonPrimary} onClick={submit} data-testid="pi-dialog-ok">OK</button>
        </div>
      </div>
    </div>
  )
}

function SelectDialog(props: {
  dialog: Extract<DialogRequest, { kind: "select" }>;
  classNames: typeof DEFAULT_CLASSNAMES;
  onResolve: (v: string | undefined) => void;
  countdown: number | null;
  onCancel: () => void;
}) {
  const { dialog, classNames: cn, onResolve, countdown, onCancel } = props;
  return (
    <div className={cn.backdrop} data-testid="pi-dialog-backdrop">
      <div className={cn.card} role="dialog" aria-modal="true">
        <h2 className={cn.title}>{dialog.title}</h2>
        {countdown !== null && <p className={cn.message} data-testid="pi-dialog-countdown">Auto-cancelling in {countdown}s</p>}
        <div className={cn.actions} style={{ flexDirection: "column", alignItems: "stretch" }}>
          {dialog.options.map((opt: string) => (
            <button key={opt} className={cn.button} onClick={() => onResolve(opt)} data-testid={`pi-dialog-option-${opt}`}>
              {opt}
            </button>
          ))}
          <button className={cn.button} onClick={onCancel} data-testid="pi-dialog-cancel">Cancel</button>
        </div>
      </div>
    </div>
  )
}

function useCountdown(timeout: number | undefined, enabled: boolean, set: (n: number | null) => void) {
  // Simple effect: tick every 1s until timeout
  if (!enabled || !timeout) {
    if (timeout) set(null)
    return
  }
  const start = Date.now()
  const tick = () => {
    const remaining = Math.max(0, Math.ceil((timeout - (Date.now() - start)) / 1000))
    set(remaining)
    if (remaining > 0) setTimeout(tick, 1000)
  }
  tick()
}

// ============================================================================
// NotificationToaster
// ============================================================================

/**
 * Renders all pending notifications as a stack. Click × to dismiss.
 */
export interface NotificationToasterProps {
  classNames?: Partial<{
    container: string;
    item: string;
    itemInfo: string;
    itemWarning: string;
    itemError: string;
    dismiss: string;
    message: string;
  }>;
}

const DEFAULT_TOASTER_CLASSNAMES = {
  container: "pi-toast-container",
  item: "pi-toast",
  itemInfo: "pi-toast pi-toast--info",
  itemWarning: "pi-toast pi-toast--warning",
  itemError: "pi-toast pi-toast--error",
  dismiss: "pi-toast-dismiss",
  message: "pi-toast-message",
};

export function NotificationToaster(props: NotificationToasterProps = {}) {
  const items = usePiNotifications();
  const adapter = useUiAdapter();
  const cn = { ...DEFAULT_TOASTER_CLASSNAMES, ...(props.classNames ?? {}) };

  if (items.length === 0) return null;

  return (
    <div className={cn.container} data-testid="pi-toast-container">
      {items.map((item) => (
        <Toast key={item.id} item={item} classNames={cn} onDismiss={() => adapter.dismissNotification(item.id)} />
      ))}
    </div>
  );
}

function Toast(props: {
  item: NotificationItem;
  classNames: typeof DEFAULT_TOASTER_CLASSNAMES;
  onDismiss: () => void;
}) {
  const { item, classNames: cn, onDismiss } = props;
  const variant =
    item.type === "warning" ? cn.itemWarning :
    item.type === "error" ? cn.itemError :
    cn.itemInfo
  return (
    <div className={variant} role="status" data-testid={`pi-toast-${item.type}`}>
      <span className={cn.message}>{item.message}</span>
      <button className={cn.dismiss} onClick={onDismiss} aria-label="Dismiss">×</button>
    </div>
  )
}

// ============================================================================
// StatusBar (optional convenience)
// ============================================================================

/** Renders the status entries as a horizontal list. */
export function PiStatusBar(props: { className?: string }) {
  const statuses = usePiStatuses();
  const entries = Array.from(statuses.entries()).filter((entry): entry is [string, string] => entry[1] !== undefined)
  if (entries.length === 0) return null
  return (
    <div className={props.className} data-testid="pi-status-bar">
      {entries.map(([key, text]) => (
        <span key={key} data-testid={`pi-status-${key}`}>{text}</span>
      ))}
    </div>
  )
}

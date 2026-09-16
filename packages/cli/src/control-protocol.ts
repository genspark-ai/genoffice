/**
 * Wire format between the genoffice CLI and the running GenOffice shell.
 *
 * The shell listens on a local socket (unix socket in userData, named pipe on
 * Windows) and publishes its endpoint plus a per-run token in
 * userData/control.json. One newline-terminated JSON request per connection,
 * one JSON reply. Both sides import this file, so the shapes stay in step.
 */

export const CONTROL_FILE = 'control.json'
export const CONTROL_PROTOCOL = 1
export const CONTROL_MAX_REQUEST_BYTES = 64 * 1024

export interface ControlEndpoint {
  protocol: number
  pid: number
  /** unix socket path or `\\.\pipe\...` name */
  endpoint: string
  token: string
}

/** Where to put the caret / selection in an open document; one kind per file type. */
export type ControlTarget =
  | { kind: 'slide'; slide: number; el?: string }
  | { kind: 'block'; block: number }
  | { kind: 'range'; sheet?: string; range: string }
  | { kind: 'page'; page: number }

export type ControlRequest =
  { cmd: 'open'; path: string; target?: ControlTarget } | { cmd: 'selection'; path: string }

export interface ControlEnvelope {
  token: string
  request: ControlRequest
}

export type ControlErrorReason =
  | 'target_not_found'
  | 'out_of_range'
  | 'sheet_not_found'
  | 'invalid_argument'
  | 'unsupported'
  | 'file_not_found'
  | 'file_not_open_in_gui'
  | 'app_unavailable'

export interface ControlError {
  reason: ControlErrorReason
  message: string
  detail?: Record<string, unknown>
}

export type ControlReply =
  { ok: true; result: Record<string, unknown> } | { ok: false; error: ControlError }

/**
 * What a renderer answers to `window.__genofficeControl(req)`; `not_ready`
 * means the document has not finished loading and the shell should ask again.
 */
export type RendererControlRequest = { cmd: 'goto'; target: ControlTarget } | { cmd: 'selection' }

export type RendererControlReply =
  | { status: 'ok'; result: Record<string, unknown> }
  | { status: 'not_ready' }
  | { status: 'error'; error: ControlError }
